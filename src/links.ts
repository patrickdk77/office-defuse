import { OfficeCategory as C, OfficeDetail as D } from './types';
import { UNKNOWN } from './xml';

export interface LinkVerdict {
  /** LINK, or MACRO for a link that runs a macro or script. */
  category: C.Link | C.Macro;
  detail: D;
  host?: string;
}

const IPV4 = /^\d{1,3}(\.\d{1,3}){3}$/;
const SCHEME = /^([a-zA-Z][a-zA-Z0-9+.-]*):/;

/** A browser drops control characters from a URL, and so does this before reading one. */
const clean = (s: string) => s.replace(/\p{Cc}/gu, '').trim();

const networkPrefix = (s: string) => s.startsWith('\\\\') || s.startsWith('//') || s.startsWith('\\/') || s.startsWith('/\\');

/** A target that names the same document or a place in it. RFC 3986 reads an empty reference as the document itself. */
export const isFragment = (target: string) => {
  const t = clean(target);
  return t === '' || t.startsWith('#');
};

/**
 * Classifies a hyperlink target with pdf-defuse's checks, without the comparison of the target with the link's text.
 * Call it only for targets that are more than a fragment. A target that holds the content of an external entity
 * cannot be read, and reads as the empty string: a relative link, which is what pdf-defuse makes of a URI it cannot
 * resolve.
 */
export function checkLink(rawTarget: string): LinkVerdict {
  const uri = rawTarget.includes(UNKNOWN) ? '' : clean(rawTarget);
  const link = (detail: D, site?: string): LinkVerdict => ({ category: C.Link, detail, host: site });
  if (networkPrefix(uri)) return link(D.NetworkPath);
  const m = SCHEME.exec(uri);
  if (!m) return link(D.Relative);
  const scheme = m[1].toLowerCase();
  // C:\path is a drive, not a scheme.
  if (scheme.length === 1) return link(D.FileUrl);
  if (scheme === 'javascript' || scheme === 'vbscript') return link(D.ScriptUrl);
  if (scheme === 'vnd.sun.star.script' || scheme === 'macro') return { category: C.Macro, detail: D.Link };
  if (scheme === 'data') return link(D.DataUrl);
  if (scheme === 'smb' || scheme === 'cifs' || scheme === 'afp' || scheme === 'nfs') return link(D.NetworkPath);
  if (scheme === 'file') {
    let host = '';
    try {
      host = new URL(uri).hostname;
    } catch {
      // A file: URL the parser refuses is read as a local path.
    }
    return link(host && host.toLowerCase() !== 'localhost' ? D.NetworkPath : D.FileUrl);
  }
  if (scheme === 'mailto') {
    let addr = uri.slice(7).split('?')[0];
    try {
      addr = decodeURIComponent(addr);
    } catch {
      // The address stays as written.
    }
    const domain = addr.includes('@') ? addr.slice(addr.lastIndexOf('@') + 1).toLowerCase() : '';
    return link(hostVerdict(domain, domain) ?? D.Safe, domain);
  }
  if (scheme !== 'http' && scheme !== 'https') return link(D.OtherScheme);
  // The authority as written, before the URL parser normalizes it.
  const rawAuth = uri
    .slice(m[0].length)
    .replace(/^[\\/]*/, '')
    .split(/[/\\?#]/, 1)[0];
  let parsed: URL;
  try {
    parsed = new URL(uri);
  } catch {
    return link(D.OtherScheme);
  }
  if (parsed.username || parsed.password || rawAuth.includes('@')) return link(D.Credentials);
  // The parser lowercases, punycodes and turns every IPv4 shorthand into dotted form, as browsers do.
  const host = parsed.hostname.replace(/\.$/, '');
  return link(hostVerdict(host, rawAuth) ?? D.Safe, host);
}

function hostVerdict(host: string, raw: string): D | undefined {
  if (!host) return D.OtherScheme;
  if (raw.includes('%')) return D.EncodedHost;
  if (host.startsWith('[') || IPV4.test(host)) return D.IpHost;
  if (/^\d+$/.test(host) || /^0x[0-9a-f]+$/i.test(host)) return D.IpHost;
  if (/\P{ASCII}/u.test(host) || host.split('.').some(label => label.startsWith('xn--'))) return D.LookalikeHost;
  return undefined;
}

/**
 * The target with every %XX turned into its character. A target can percent-encode its scheme, as `%73cript:` does in
 * a CVE-2017-0199 sample, so the checks below hold a target local or inside the package only when both forms are.
 */
const decoded = (s: string) => s.replace(/%([0-9a-fA-F]{2})/g, (_, hex: string) => String.fromCharCode(Number.parseInt(hex, 16)));

function local(t: string): boolean {
  if (networkPrefix(t)) return false;
  const m = SCHEME.exec(t);
  if (!m || m[1].length === 1) return true;
  if (m[1].toLowerCase() !== 'file') return false;
  try {
    const host = new URL(t).hostname;
    return !host || host.toLowerCase() === 'localhost';
  } catch {
    return true;
  }
}

/**
 * Whether a target outside the file is on this machine: a drive path, a file: URL without a host, an absolute path or
 * a relative one. A UNC path, a file: URL with a host and every other scheme reach elsewhere, and so can a target that
 * holds the content of an external entity.
 */
export const isLocalTarget = (rawTarget: string) => !rawTarget.includes(UNKNOWN) && local(clean(rawTarget)) && local(clean(decoded(rawTarget)));

function inside(t: string): boolean {
  if (t === '' || networkPrefix(t) || SCHEME.test(t) || t.startsWith('/') || t.startsWith('\\')) return false;
  let depth = 0;
  for (const seg of t.split(/[/\\]/)) {
    if (seg === '..') depth--;
    else if (seg !== '.' && seg !== '') depth++;
    if (depth < 0) return false;
  }
  return true;
}

/** Whether an ODF reference names an entry inside the package: a relative path that does not climb out of it. */
export const isInsidePackage = (rawHref: string) => inside(clean(rawHref)) && inside(clean(decoded(rawHref)));

/** Whether an OOXML relationship target is not a part name, whatever TargetMode says: a scheme, a drive or a network path. */
export const isAbsoluteTarget = (rawTarget: string) => [rawTarget, decoded(rawTarget)].some(t => SCHEME.test(clean(t)) || t.startsWith('\\') || t.startsWith('//'));
