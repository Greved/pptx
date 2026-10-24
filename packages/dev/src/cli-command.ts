import { accessSync, closeSync, constants, openSync, readSync, statSync } from 'node:fs';
import { basename, delimiter, resolve } from 'node:path';

interface CliCommand {
  file: string;
  args: string[];
}

const SHEBANG_LIMIT = 4096;

const lookupFailure = (cause: unknown): boolean =>
  cause instanceof Error &&
  'code' in cause &&
  (cause.code === 'ENOENT' || cause.code === 'ENOTDIR' || cause.code === 'EACCES');

const hasWindowsLauncher = (directory: string, command: string): boolean =>
  ['.com', '.exe', '.cmd', '.bat'].some((extension) => {
    try {
      return statSync(resolve(directory, command + extension)).isFile();
    } catch (cause) {
      if (!lookupFailure(cause)) throw cause;
      return false;
    }
  });

/** Resolve Node-shebang CLIs without introducing a command shell. */
export const resolveCliCommand = (
  command: string,
  env: NodeJS.ProcessEnv = process.env,
  cwd: string = process.cwd(),
): CliCommand => {
  const unchanged = { file: command, args: [] };
  if (basename(command) !== command) return unchanged;
  if (process.platform === 'win32' && hasWindowsLauncher(cwd, command)) return unchanged;
  const path =
    process.platform === 'win32'
      ? Object.entries(env).find(([key]) => key.toUpperCase() === 'PATH')?.[1]
      : env.PATH;
  if (path === undefined) return unchanged;
  for (const directory of path.split(delimiter)) {
    const location = resolve(cwd, directory);
    // Keep native/shim lookup ahead of a Node script later in PATH, including
    // the Windows current-directory lookup that precedes PATH.
    if (process.platform === 'win32' && hasWindowsLauncher(location, command)) return unchanged;
    const file = resolve(location, command);
    let descriptor: number | undefined;
    try {
      if (!statSync(file).isFile()) continue;
      // Invoking Node directly must not make a non-executable POSIX file runnable.
      if (process.platform !== 'win32') accessSync(file, constants.X_OK);
      descriptor = openSync(file, 'r');
      const buffer = Buffer.alloc(SHEBANG_LIMIT);
      const length = readSync(descriptor, buffer, 0, buffer.length, 0);
      const content = buffer.toString('utf8', 0, length);
      const newline = content.indexOf('\n');
      if (newline < 0 && length === SHEBANG_LIMIT) return unchanged;
      const shebang = (newline < 0 ? content : content.slice(0, newline)).trimEnd();
      // Windows cannot execute shebang files. Recognize only these Node forms;
      // native executables and other interpreters retain their normal spawn path.
      if (shebang === `#!${process.execPath}` || /^#!\/usr\/bin\/env[\t ]+node$/.test(shebang))
        return { file: process.execPath, args: [file] };
      return unchanged;
    } catch (cause) {
      if (!lookupFailure(cause)) throw cause;
    } finally {
      if (descriptor !== undefined) closeSync(descriptor);
    }
  }
  return unchanged;
};
