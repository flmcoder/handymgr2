import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';

test('vacancies list resolves properties through the numeric AppFolio link id', async () => {
  const source = await readFile(new URL('./server.ts', import.meta.url), 'utf8');

  assert.match(source, /from appfolio_unit_vacancies v\s*\n\s*join appfolio_properties p on p\.raw_json->>'Link'/);
  assert.match(source, /PropertyGroupIds'\s*@>\s*jsonb_build_array/);
});

test('vacancies response prefers the stored days_vacant count', async () => {
  const source = await readFile(new URL('./server.ts', import.meta.url), 'utf8');

  assert.match(source, /storedRaw/);
  assert.match(source, /row\.days_vacant/);
});

test('property performance resolves vacancy counts through the numeric link', async () => {
  const source = await readFile(new URL('./server.ts', import.meta.url), 'utf8');

  assert.match(source, /join appfolio_properties p2 on p2\.raw_json->>'Link'/);
});
