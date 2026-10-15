/** Splits a migration file into executable statements, dropping `--` comment lines. */
export function splitSqlStatements(rawSql: string): string[] {
  return String(rawSql || '')
    .split('\n')
    .filter((line) => !line.trim().startsWith('--'))
    .join('\n')
    .split(/;\s*(?:\n|$)/)
    .map((statement) => statement.trim())
    .filter(Boolean);
}
