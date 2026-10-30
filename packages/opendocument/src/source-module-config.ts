/** Read the scalar config an OpenDocument source module receives (fail closed). */
export function readOdfSourceModuleConfig(config: unknown, label: string): Readonly<{ wasmUrl: string }> {
  if (typeof config !== 'object' || config === null) {
    throw new TypeError(`${label} source module config must be an object`);
  }
  const record = config as Record<string, unknown>;
  for (const key of Object.keys(record)) {
    if (key !== 'wasmUrl') throw new TypeError(`${label} source module config has an unknown field`);
  }
  if (typeof record.wasmUrl !== 'string') {
    throw new TypeError(`${label} wasmUrl must be a nonempty absolute URL`);
  }
  return { wasmUrl: record.wasmUrl };
}
