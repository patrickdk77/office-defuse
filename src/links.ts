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

/** A host as LabelSites compares it. */
const siteName = (h: string) =>
  h
    .toLowerCase()
    .replace(/^www\./, '')
    .replace(/\.$/, '');

/**
 * The hosts a label names, as a lookup: has() answers whether a host is one of them or a subdomain of one, or one of
 * them is a subdomain of the host. No public-suffix list is needed: evil.co.uk is not under mybank.co.uk. The same
 * comparison as pdf-defuse's.
 */
export class LabelSites {
  private readonly hosts = new Set<string>();
  /** Every domain some label host is a subdomain of. */
  private readonly parents = new Set<string>();
  constructor(names: string[]) {
    for (const n of names) {
      const x = siteName(n);
      this.hosts.add(x);
      for (let i = x.indexOf('.'); i >= 0; i = x.indexOf('.', i + 1)) this.parents.add(x.slice(i + 1));
    }
  }
  get size(): number {
    return this.hosts.size;
  }
  has(host: string): boolean {
    const y = siteName(host);
    if (this.hosts.has(y) || this.parents.has(y)) return true;
    for (let i = y.indexOf('.'); i >= 0; i = y.indexOf('.', i + 1)) if (this.hosts.has(y.slice(i + 1))) return true;
    return false;
  }
}

/**
 * Top-level domains a bare host in a label may end in: the common generic ones and the country codes. A name that
 * ends in anything else, such as an Excel function (BETA.DIST) or a file (moon.htm), is not read as an address.
 */
const TLDS = new Set(
  (
    'com net org edu gov mil int info biz name pro mobi asia tel travel jobs museum aero coop app dev io ai co me tv cc ly ' +
    'xyz online site top shop club store tech live page blog cloud email link news ' +
    'ac ad ae af ag ai al am ao aq ar as at au aw ax az ba bb bd be bf bg bh bi bj bm bn bo br bs bt bw by bz ca cd cf cg ' +
    'ch ci ck cl cm cn cr cu cv cw cx cy cz de dj dk dm do dz ec ee eg er es et eu fi fj fk fm fo fr ga gd ge gf gg gh gi ' +
    'gl gm gn gp gq gr gs gt gu gw gy hk hm hn hr ht hu id ie il im in iq ir is it je jm jo jp ke kg kh ki km kn kp kr kw ' +
    'ky kz la lb lc li lk lr ls lt lu lv ma mc md mg mh mk ml mm mn mo mp mq mr ms mt mu mv mw mx my mz na nc ne nf ng ni ' +
    'nl no np nr nu nz om pa pe pf pg ph pk pl pm pn pr ps pt pw py qa re ro rs ru rw sa sb sc sd se sg sh si sk sl sm sn ' +
    'so sr ss st su sv sx sy sz tc td tf tg th tj tk tl tm tn to tr tt tw tz ua ug uk us uy uz va vc ve vg vi vn vu wf ws ' +
    'ye yt za zm zw'
  ).split(' '),
);

/** The most of a label that is read. */
const MAX_LABEL = 4096;

/**
 * The site a label names when the label is itself an address: a web URL, a mail address, or a host with or without a
 * path that starts with www. or ends in a known top-level domain. Text that only mentions an address, an Excel
 * function name and a file name name no site. Excel's function names are all capitals, as typed hosts never are, so
 * a bare host in capitals counts only after WWW.
 */
export function labelSite(raw: string): string | undefined {
  const t = raw
    .slice(0, MAX_LABEL)
    .replace(/\p{Cc}/gu, '')
    .trim();
  if (/^https?:\/\//i.test(t)) {
    try {
      return new URL(t).hostname.toLowerCase() || undefined;
    } catch {
      return undefined;
    }
  }
  const mail = /^(?:mailto:)?[^\s@/]+@([a-z0-9-]+(?:\.[a-z0-9-]+)+)$/i.exec(t);
  const host = mail?.[1] ?? /^([a-z0-9-]+(?:\.[a-z0-9-]+)+)(?::\d+)?(?:[/?#]\S*)?$/i.exec(t)?.[1];
  if (!host) return undefined;
  const h = host.toLowerCase();
  if (h.startsWith('www.')) return h;
  if (!mail && host === host.toUpperCase() && /[A-Z]/.test(host)) return undefined;
  return TLDS.has(h.slice(h.lastIndexOf('.') + 1)) ? h : undefined;
}

/**
 * The first label of a safe link that is an address, and an address of another site than the one the link opens, or
 * undefined. A label is text the reader sees for the link: its text, its tooltip, a display name. A subdomain or a
 * parent of the label's site counts as the same site, as in pdf-defuse. A target that fails on its own gets its own
 * finding instead.
 */
export function mismatchedLabel(target: string, labels: ReadonlyArray<string | undefined>): string | undefined {
  if (!labels.some(Boolean)) return undefined;
  const v = checkLink(target);
  if (v.category !== C.Link || v.detail !== D.Safe || !v.host) return undefined;
  for (const label of labels) {
    const site = label ? labelSite(label) : undefined;
    // A site without a dot, where the target has one, is a label cut short or a local alias, and names no site.
    if (!site || (!site.includes('.') && v.host.includes('.'))) continue;
    if (!new LabelSites([site]).has(v.host)) return label;
  }
  return undefined;
}
