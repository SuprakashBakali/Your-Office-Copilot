import { ensureLockdown } from "./lockdown";

/* global Compartment */

export function sandboxedEval(
  code: string,
  globals: Record<string, unknown>,
): unknown {
  ensureLockdown();
  
  // Create a new compartment with restricted globals
  // @ts-ignore - Compartment comes from ses
  const compartment = new Compartment({
    globals: {
      ...globals,
      console,
      Math,
      Date,
      atob: atob.bind(globalThis),
      btoa: btoa.bind(globalThis),
      Function: undefined,
      Reflect: undefined,
      Proxy: undefined,
      Compartment: undefined,
      harden: undefined,
      lockdown: undefined,
    },
    __options__: true,
  });
  
  // Wrap in async IIFE so we can use await inside the eval block
  return compartment.evaluate(`(async () => { ${code} })()`);
}
