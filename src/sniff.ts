/**
 * Identifies a file type from its first bytes, with the MIME types pdf-defuse gives, so plugins of the family see one
 * name for each type. Returns a MIME type or undefined.
 */
export function sniffType(head: Uint8Array): string | undefined {
  const b = Buffer.from(head.buffer, head.byteOffset, head.byteLength);
  // PDF readers accept the header anywhere in the first 1024 bytes.
  const s = b.subarray(0, 1024).toString('latin1');
  if (s.includes('%PDF-')) return 'application/pdf';
  if (b[0] === 0x50 && b[1] === 0x4b && (b[2] === 0x03 || b[2] === 0x05 || b[2] === 0x07)) return 'application/zip';
  if (b[0] === 0x89 && s.startsWith('\x89PNG')) return 'image/png';
  if (b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return 'image/jpeg';
  if (s.startsWith('GIF87a') || s.startsWith('GIF89a')) return 'image/gif';
  if (b[0] === 0xd0 && b[1] === 0xcf && b[2] === 0x11 && b[3] === 0xe0) return 'application/x-ole-storage';
  if (s.startsWith('{\\rt')) return 'application/rtf';
  if (s.startsWith('MZ')) return 'application/x-msdownload';
  if (s.startsWith('\x7fELF')) return 'application/x-executable';
  const text = s.replace(/^\xef\xbb\xbf/, '').trimStart();
  if (/^<svg[\s>/]/i.test(text) || (/^<\?xml/i.test(text) && /<svg[\s>/]/i.test(s))) return 'image/svg+xml';
  if (/^<\?xml/i.test(text)) return 'application/xml';
  if (/^<!doctype html|^<html[\s>]/i.test(text)) return 'text/html';
  return undefined;
}
