import { InMemoryFs, Bash } from "just-bash/browser";

// Initialize a persistent VFS and Bash shell for the current session
export const vfs = new InMemoryFs();
export const bash = new Bash({ fs: vfs });

// Ensure the home directory exists
try {
  vfs.mkdirSync("/home/user", { recursive: true });
} catch (e) {
  // Ignore if already exists
}

/**
 * Helper to run a bash command and return its stdout and stderr combined
 */
export async function runBashCommand(command: string): Promise<string> {
  const result = await bash.exec(command);
  
  const stdout = result.stdout || '';
  const stderr = result.stderr || '';
  
  if (result.exitCode !== 0) {
    return `Error (Exit Code ${result.exitCode}):\n${stderr}\n${stdout}`.trim();
  }
  
  return (stdout + '\n' + stderr).trim() || 'Command executed successfully (no output).';
}
