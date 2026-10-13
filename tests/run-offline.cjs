// Exercise the production scanner and result conversion with Windows Script Host
// and disconnected ADO recordsets, without changing Excel's VBA trust settings.
// WSH has no typed declarations, Like operator, labels or one-based arrays. The
// narrow adaptations below preserve the tested paths; Excel tests cover bounds,
// formula evaluation, caching and connection isolation using the actual VBA.
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {execFileSync} = require('node:child_process');
const root = path.resolve(__dirname, '..');
function common(source) {
  return source.split(/\r?\n/).filter(line => !/^Attribute |^Option Explicit/.test(line)).map(line => {
    if (/^\s*(?:(?:Public|Private)\s+)?(?:Function|Sub|Dim|Const)\b/.test(line)) {
      line = line.replace(/\b(?:Public|Private)\s+/g, '').replace(/\s+As\s+\w+/g, '').replace(/\bOptional\s+/g, '').replace(/onlyColumn = ""/, 'onlyColumn');
    }
    return line.replace(/\b(Mid|Left|Right|Chr|UCase|LCase|Trim|Format)\$/g, '$1').replace(/^(\s*Next)\s+\w+\s*$/, '$1').replace(/\b(\d+)#/g, '$1');
  }).join('\n');
}
let sql = common(fs.readFileSync(path.join(root, 'src/SQLValidation.bas'), 'utf8'));
for (const label of ['NotSelect', 'Malformed']) {
  const body = sql.match(new RegExp(label + ':\\n\\s*(ValidateSelect = [^\\n]+)'))?.[1];
  if (!body) throw new Error('Scanner label adaptation failed: ' + label);
  sql = sql.replaceAll('GoTo ' + label, body + ': Exit Function');
}
sql = sql.replace(/NotSelect:[\s\S]*?\nEnd Function/, 'End Function');
sql = sql.replace('ch Like "[A-Za-z_]"', 'CharMatch(ch, "^[A-Za-z_]$")')
  .replace('Mid(sql, i, 1) Like "[A-Za-z0-9_$]"', 'CharMatch(Mid(sql, i, 1), "^[A-Za-z0-9_$]$")');
let results = common(fs.readFileSync(path.join(root, 'src/FormulaResults.bas'), 'utf8'));
const overflow = results.match(/TooLarge:\n\s*(Err.Raise [^\n]+)/)?.[1];
if (!overflow) throw new Error('Result label adaptation failed');
results = results.replaceAll('GoTo TooLarge', overflow).replace(/TooLarge:\n\s*Err.Raise [^\n]+\n/, '');
results = results.replace('output(1 To rowCount + headerRows, 1 To colCount)', 'output(rowCount + headerRows, colCount)');
const helper = `
Function CharMatch(value, pattern)
  Dim expression
  Set expression = New RegExp
  expression.Pattern = pattern
  CharMatch = expression.Test(value)
End Function
`;
const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'Excelerator-tests-'));
try {
  const script = path.join(temporary, 'offline.vbs');
  fs.writeFileSync(script, 'Option Explicit\n' + sql + '\n' + results + helper + fs.readFileSync(path.join(__dirname, 'offline.vbs'), 'utf8'));
  process.stdout.write(execFileSync('cscript.exe', ['//nologo', script], {encoding:'utf8', timeout:60000}));
} catch (error) {
  process.stderr.write(error.stdout || '');
  process.stderr.write(error.stderr || '');
  const line = /offline\.vbs\((\d+),/.exec(error.stderr || '')?.[1];
  if (line) {
    const source = fs.readFileSync(path.join(temporary, 'offline.vbs'), 'utf8').split('\n');
    process.stderr.write(source.slice(Math.max(0, Number(line)-3), Number(line)+2).map((text, index) => `${Math.max(1, Number(line)-2)+index}: ${text}`).join('\n') + '\n');
  }
  throw error;
} finally {
  fs.rmSync(temporary, {recursive:true, force:true});
}
