declare module 'unzipper' {
  import type { Readable } from 'node:stream';

  export interface ZipEntry {
    path: string;
    type: 'File' | 'Directory' | string;
    buffer(): Promise<Buffer>;
    stream(): Readable;
  }

  export interface ZipDirectory {
    files: ZipEntry[];
  }

  export const Open: {
    file(filePath: string): Promise<ZipDirectory>;
  };

  const unzipper: {
    Open: typeof Open;
  };
  export default unzipper;
}
