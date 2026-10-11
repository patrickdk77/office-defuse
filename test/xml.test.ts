import { expect } from 'chai';
import { StopError } from '../src/errors';
import { isNcName, showReferences, UNKNOWN, unknownEntity, type XmlDoctype, type XmlElement, type XmlExternalReference, type XmlLimits, type XmlOptions, XmlTokenizer } from '../src/xml';

interface Event {
  e: 'start' | 'end' | 'text';
  v: string;
}

const LIMITS: XmlLimits = { depth: 256, tokenSize: 1 << 24, entityDepth: 16, entityExpansions: 100_000, entitySize: 1 << 24 };

/** Tokenizes `input`, fed in pieces of `piece` bytes, and returns the events with adjacent text joined. */
function events(input: string | Uint8Array, piece = 1 << 20, limits = LIMITS, options: XmlOptions = {}): Event[] {
  const out: Event[] = [];
  const t = new XmlTokenizer(
    {
      start: (el: XmlElement) => {
        const attrs = el.attributes.map(a => ` {${a.ns}}${a.local}=${a.value}`).join('');
        out.push({ e: 'start', v: `{${el.ns}}${el.local}${attrs}` });
      },
      end: (el: XmlElement) => out.push({ e: 'end', v: `{${el.ns}}${el.local}` }),
      text: (s: string) => {
        const last = out[out.length - 1];
        if (last?.e === 'text') last.v += s;
        else out.push({ e: 'text', v: s });
      },
    },
    limits,
    'test.xml',
    options,
  );
  const bytes = typeof input === 'string' ? Buffer.from(input, 'utf8') : input;
  for (let i = 0; i < bytes.length; i += piece) t.write(bytes.subarray(i, i + piece));
  t.end();
  return out;
}

/** What the tokenizer refuses `input` with, as category/detail, or 'accepted'. */
function refusal(input: string | Uint8Array, limits = LIMITS, piece = 1 << 20, options: XmlOptions = {}): string {
  try {
    events(input, piece, limits, options);
  } catch (e) {
    if (e instanceof StopError) return `${e.category}/${e.detail}`;
    throw e;
  }
  return 'accepted';
}

/** The reason the tokenizer gives for refusing `input`. */
function reason(input: string, limits = LIMITS): string {
  try {
    events(input, 1 << 20, limits);
  } catch (e) {
    if (e instanceof StopError) return String(e.data?.reason);
    throw e;
  }
  return 'accepted';
}

/** The events, document type declarations and external entity references of `input`. */
function noted(input: string, piece = 1 << 20) {
  const doctypes: XmlDoctype[] = [];
  const references: XmlExternalReference[] = [];
  const ev = events(input, piece, LIMITS, { notes: { doctype: d => doctypes.push(d), externalEntity: r => references.push(r) } });
  return { ev, doctypes, references };
}

/** The attributes of each element of `input`, by qname, in document order. */
function attributesOf(input: string): Array<Record<string, string>> {
  const out: Array<Record<string, string>> = [];
  const t = new XmlTokenizer({ start: el => out.push(Object.fromEntries(el.attributes.map(a => [a.qname, a.value]))) }, LIMITS, 'test.xml');
  t.write(Buffer.from(input, 'utf8'));
  t.end();
  return out;
}

/** A chain of entities e1 to e`n`, each referring to the next, and the last to text. */
const chain = (n: number) => Array.from({ length: n }, (_, i) => `<!ENTITY e${i + 1} "${i + 1 < n ? `&e${i + 2};` : 'end'}">`).join('');

/** The billion laughs: ten levels of ten references each. */
const LAUGHS = `<!DOCTYPE lolz [<!ENTITY lol "lol">${Array.from({ length: 9 }, (_, i) => `<!ENTITY lol${i + 1} "${`&lol${i || ''};`.repeat(10)}">`).join('')}]><lolz>&lol9;</lolz>`;

const utf16le = (s: string) => Buffer.from(s, 'utf16le');
const utf16be = (s: string) => Buffer.from(s, 'utf16le').swap16();

describe('XML tokenizer', () => {
  it('resolves namespaces by URI, whatever the prefix', () => {
    const ev = events('<?xml version="1.0"?><x:doc xmlns:x="urn:a" xmlns="urn:b"><y x:k="1" k="2"/></x:doc>');
    expect(ev).to.deep.equal([
      { e: 'start', v: '{urn:a}doc' },
      { e: 'start', v: '{urn:b}y {urn:a}k=1 {}k=2' },
      { e: 'end', v: '{urn:b}y' },
      { e: 'end', v: '{urn:a}doc' },
    ]);
  });

  it('binds the xml prefix without a declaration', () => {
    expect(events('<a xml:space="preserve"/>')[0].v).to.equal('{}a {http://www.w3.org/XML/1998/namespace}space=preserve');
  });

  it('expands the predefined entities and character references in text and attributes', () => {
    const ev = events('<a v="D&#68;E&#x41;UTO &lt;&amp;&gt;">x&#x1F600;&apos;&quot;</a>');
    expect(ev[0].v).to.equal('{}a {}v=DDEAUTO <&>');
    expect(ev[1].v).to.equal(`x${String.fromCodePoint(0x1f600)}'"`);
  });

  it('normalizes line ends in text and white space in attribute values, keeping references', () => {
    const ev = events('<a v="1\r\n2\t3&#10;4">x\r\ny\rz</a>');
    expect(ev[0].v).to.equal('{}a {}v=1 2 3\n4');
    expect(ev[1].v).to.equal('x\ny\nz');
  });

  it('reads CDATA as text, skips comments and processing instructions, and takes single quotes and spaces around "="', () => {
    const ev = events("<?xml version='1.0' encoding='UTF-8'?><!-- c --><?pi data?><?pi?><a b = '1' ><![CDATA[<b>&amp;]]><!-- <b/> --></a >");
    expect(ev).to.deep.equal([
      { e: 'start', v: '{}a {}b=1' },
      { e: 'text', v: '<b>&amp;' },
      { e: 'end', v: '{}a' },
    ]);
  });

  it('gives the same events however the bytes are split', () => {
    const doc =
      '<?xml version="1.0" encoding="UTF-8"?>\r\n<r xmlns:p="urn:p"><p:a p:x="caf&#233; &amp; t\xe9">t\xe9xt &lt; &#x10000;\r\nmore<![CDATA[ x\r\n ]]></p:a><!--c--><b/><?p q?><longer-element-name></longer-element-name></r>';
    const whole = events(doc);
    for (const piece of [1, 2, 3, 7]) expect(events(doc, piece), `piece ${piece}`).to.deep.equal(whole);
  });

  it('reads UTF-16 with a byte order mark in either order, UTF-16 announced without one, and UTF-8 with one', () => {
    const doc = '<?xml version="1.0" encoding="UTF-16"?><a>\xe9t\xe9</a>';
    const expected = [
      { e: 'start', v: '{}a' },
      { e: 'text', v: '\xe9t\xe9' },
      { e: 'end', v: '{}a' },
    ];
    for (const b of [Buffer.concat([Buffer.from([0xff, 0xfe]), utf16le(doc)]), Buffer.concat([Buffer.from([0xfe, 0xff]), utf16be(doc)]), utf16le(doc), utf16be(doc)])
      expect(events(b, 3)).to.deep.equal(expected);
    expect(events(Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from('<a/>')]))).to.have.length(2);
    // A part shorter than four bytes is still read.
    expect(events('<a/>')).to.have.length(2);
    expect(refusal('<a>')).to.equal('XML/MALFORMED');
  });

  it('reports a document type declaration, with its identifiers and offsets, and loads nothing it names', () => {
    const manifest =
      '<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE manifest:manifest PUBLIC "-//OpenOffice.org//DTD Manifest 1.0//EN" "Manifest.dtd">\n<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">\n <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.text" manifest:full-path="/"/>\n</manifest:manifest>';
    const whole = noted(manifest);
    expect(whole.doctypes).to.have.length(1);
    const d = whole.doctypes[0];
    expect(d).to.deep.include({ name: 'manifest:manifest', publicId: '-//OpenOffice.org//DTD Manifest 1.0//EN', systemId: 'Manifest.dtd', internalSubset: false, entities: 0, attributeDefaults: 0 });
    expect(manifest.slice(d.start, d.end)).to.match(/^<!DOCTYPE manifest:manifest PUBLIC .*"Manifest.dtd">$/);
    expect(whole.ev.map(e => e.e)).to.deep.equal(['start', 'text', 'start', 'end', 'text', 'end']);
    // A declaration that arrives in pieces reads the same, with or without an internal subset.
    for (const piece of [1, 2, 3, 7]) expect(noted(manifest, piece), `piece ${piece}`).to.deep.equal(whole);
    const subset = '<!DOCTYPE r [ <!ENTITY e "x>"> <!-- ]> --> <?p ]>?> <!ATTLIST r a CDATA \'>\'> ]><r>&e;</r>';
    const read = noted(subset);
    expect(read.ev).to.deep.equal([
      { e: 'start', v: '{}r {}a=>' },
      { e: 'text', v: 'x>' },
      { e: 'end', v: '{}r' },
    ]);
    for (const piece of [1, 2, 3, 5]) expect(noted(subset, piece), `piece ${piece}`).to.deep.equal(read);
    expect(noted('<!DOCTYPE a SYSTEM "a.dtd"><a/>').doctypes[0]).to.deep.include({ name: 'a', systemId: 'a.dtd', internalSubset: false });
    expect(noted('<!DOCTYPE a><a/>').doctypes[0]).to.deep.include({ name: 'a', internalSubset: false });
    expect(noted('<!DOCTYPE a[]><a/>').doctypes[0]).to.deep.include({ name: 'a', internalSubset: true });
    // The external subset is never read, so an entity only it could declare stays undeclared.
    expect(reason('<!DOCTYPE a SYSTEM "a.dtd"><a>&nbsp;</a>')).to.equal('A reference to entity "nbsp", which nothing declares');
  });

  it('expands internal entities in text and attribute values, markup included, and reports the counts', () => {
    const doc = '<!DOCTYPE r [\r\n<!ENTITY e "D&#68;E">\r\n<!ENTITY m "<b k=\'&#38;amp;&#38;e;\'>&e;\r\n</b>">\r\n<!ENTITY x SYSTEM "x.xml">\r\n]><r v="&e;AUTO">&e;AUTO&m;</r>';
    const { ev, doctypes } = noted(doc);
    expect(ev).to.deep.equal([
      { e: 'start', v: '{}r {}v=DDEAUTO' },
      { e: 'text', v: 'DDEAUTO' },
      { e: 'start', v: '{}b {}k=&DDE' },
      { e: 'text', v: 'DDE\n' },
      { e: 'end', v: '{}b' },
      { e: 'end', v: '{}r' },
    ]);
    expect(doctypes[0]).to.deep.include({ name: 'r', internalSubset: true, entities: 3, externalEntities: 1, skipped: 0 });
    // A white space character in entity text is a space in an attribute value. A CR from a character reference stays a
    // character of its own in text, and is a space of its own in a value.
    expect(attributesOf('<!DOCTYPE r [<!ENTITY s "a&#13;&#10;b\tc">]><r v="&s;"/>')[0].v).to.equal('a  b c');
    expect(events('<!DOCTYPE r [<!ENTITY s "a&#13;&#10;b">]><r>&s;</r>')[1].v).to.equal('a\r\nb');
    // The first declaration of a name binds, and the predefined entities keep their meaning.
    expect(events('<!DOCTYPE r [<!ENTITY t "one"><!ENTITY t "two"><!ENTITY lt "x">]><r>&t;&lt;</r>')[1].v).to.equal('one<');
    // Elements from entity text carry the offsets of the outermost reference.
    const nested = '<!DOCTYPE r [<!ENTITY i "<i/>"><!ENTITY o "<o>&i;</o>">]><r>&o;</r>';
    const seen: string[] = [];
    const t = new XmlTokenizer({ start: el => seen.push(`${el.local} ${nested.slice(el.start, el.end)}`) }, LIMITS, 'x');
    t.write(Buffer.from(nested));
    t.end();
    expect(seen).to.deep.equal(['r <r>', 'o &o;', 'i &o;']);
    // A CDATA section, a comment and a processing instruction in entity text are read as in the part.
    expect(events('<!DOCTYPE r [<!ENTITY c "<![CDATA[&#38;x]]><!--c--><?p q?>">]><r>&c;</r>')[1].v).to.equal('&x');
  });

  it('reports an external entity in text without fetching it, and keeps its place', () => {
    const doc = '<!DOCTYPE a [<!ENTITY e PUBLIC "-//X//Y" "http://example.invalid/e.xml"><!ENTITY i "[&e;]">]><a>x&e;y&i;</a>';
    const { ev, references } = noted(doc);
    expect(ev[1].v).to.equal(`x${UNKNOWN}e${UNKNOWN}y[${UNKNOWN}e${UNKNOWN}]`);
    expect(showReferences(ev[1].v)).to.equal('x&e;y[&e;]');
    expect(showReferences('plain')).to.equal('plain');
    expect(unknownEntity(ev[1].v)).to.equal('e');
    expect(unknownEntity('plain')).to.equal(undefined);
    const id = { name: 'e', publicId: '-//X//Y', systemId: 'http://example.invalid/e.xml' };
    const direct = doc.indexOf('x&e;y') + 1;
    const outer = doc.indexOf('&i;');
    expect(references).to.deep.equal([
      { ...id, start: direct, end: direct + 3 },
      { ...id, start: outer, end: outer + 3 },
    ]);
    expect(noted('<!DOCTYPE a [<!ENTITY e SYSTEM "e.xml">]><a>&e;</a>').references[0]).to.not.have.property('publicId');
  });

  it('refuses references XML 1.0 refuses', () => {
    const cases: Array<[string, string]> = [
      ['<a>&u;</a>', 'A reference to entity "u", which nothing declares'],
      ['<!DOCTYPE a><a v="&u;"/>', 'A reference to entity "u", which nothing declares'],
      ['<!DOCTYPE a [<!ENTITY e SYSTEM "x.txt">]><a v="&e;"/>', 'A reference to external entity "e" in an attribute value'],
      ['<!DOCTYPE a [<!ENTITY e SYSTEM "x.txt"><!ENTITY i "&e;">]><a v="&i;"/>', 'A reference to external entity "e" in an attribute value'],
      ['<!DOCTYPE a [<!ENTITY e "&#60;">]><a v="&e;"/>', '"<" in an attribute value'],
      ['<!DOCTYPE r [<!ENTITY a "&b;"><!ENTITY b "&a;">]><r>&a;</r>', 'Entity "a" refers to itself'],
      ['<!DOCTYPE r [<!ENTITY a "&a;">]><r v="&a;"/>', 'Entity "a" refers to itself'],
      ['<!DOCTYPE r [<!ENTITY e "<b>">]><r>&e;</b></r>', 'The text of entity "e" leaves <b> open'],
      ['<!DOCTYPE r [<!ENTITY e "</r>">]><r>&e;', '</r> closes an element from outside the entity text it stands in'],
      ['<!DOCTYPE r [<!ENTITY e "<b">]><r>&e;</r>', 'Entity text ends inside markup'],
      ['<!DOCTYPE r [<!ENTITY e "x">]>&e;<r/>', 'An entity reference outside the root element'],
      ['<!DOCTYPE r [<!NOTATION n SYSTEM "n"><!ENTITY u SYSTEM "u.bin" NDATA n>]><r>&u;</r>', 'A reference to entity "u", whose data is not XML'],
      ['<!DOCTYPE r [<!ENTITY e "<?xml version=\'1.0\'?>">]><r>&e;</r>', 'An XML declaration that is not at the start'],
      ['<!DOCTYPE r [<!ENTITY e "<!DOCTYPE r>">]><r>&e;</r>', 'A document type declaration after the root element, or a second one'],
    ];
    for (const [doc, why] of cases) expect(reason(doc), doc).to.equal(why);
  });

  it('refuses document type declarations that are not well-formed, and declarations outside one', () => {
    for (const doc of [
      '<a/><!DOCTYPE a>',
      '<!DOCTYPE a><!DOCTYPE a><a/>',
      '<a><!ENTITY x "y"></a>',
      '<a><!ELEMENT a ANY></a>',
      '<a><!ATTLIST a b CDATA #IMPLIED></a>',
      '<a><!NOTATION n SYSTEM "x"></a>',
      '<!doctype a><a/>',
      '<!DOCTYPEa><a/>',
      '<!DOCTYPE 1a><a/>',
      '<!DOCTYPE a SYSTEM><a/>',
      '<!DOCTYPE a SYSTEM x><a/>',
      '<!DOCTYPE a PUBLIC "{" "x"><a/>',
      '<!DOCTYPE a PUBLIC "p"><a/>',
      '<!DOCTYPE a PUBLIC "p""x"><a/>',
      '<!DOCTYPE a OTHER "x"><a/>',
      '<!DOCTYPE a [] x><a/>',
      '<!DOCTYPE a [<!FOO>]><a/>',
      '<!DOCTYPE a [ " ]><a/>',
      '<!DOCTYPE a [<![INCLUDE[<!ENTITY x "y">]]>]><a/>',
      '<!DOCTYPE a [<!ENTITY>]><a/>',
      '<!DOCTYPE a [<!ENTITY x>]><a/>',
      '<!DOCTYPE a [<!ENTITY 1x "y">]><a/>',
      '<!DOCTYPE a [<!ENTITY a:x "y">]><a/>',
      '<!DOCTYPE a [<!ENTITY x "y" z>]><a/>',
      '<!DOCTYPE a [<!ENTITY x y>]><a/>',
      '<!DOCTYPE a [<!ENTITY x SYSTEM "y" NDATA>]><a/>',
      '<!DOCTYPE a [<!ENTITY x SYSTEM "y" NDATA 1n>]><a/>',
      '<!DOCTYPE a [<!ENTITY % x SYSTEM "y" NDATA n>]><a/>',
      '<!DOCTYPE a [<!ENTITY %x "y">]><a/>',
      '<!DOCTYPE a [<!ENTITY x "%y;">]><a/>',
      '<!DOCTYPE a [<!ENTITY x "&#0;">]><a/>',
      '<!DOCTYPE a [<!ENTITY x "&1;">]><a/>',
      '<!DOCTYPE a [<!ENTITY x "&y">]><a/>',
      '<!DOCTYPE a [<!ELEMENT a %e;>]><a/>',
      '<!DOCTYPE a [<!ELEMENTa>]><a/>',
      '<!DOCTYPE a [<?xml x?>]><a/>',
      '<!DOCTYPE a [<?1 x?>]><a/>',
      '<!DOCTYPE a [%e]><a/>',
      '<!DOCTYPE a [%1;]><a/>',
      '<!DOCTYPE a [<!ENTITY % n "t"><!ENTITY %n; "x">]><a/>',
      '<!DOCTYPE a [<!ATTLIST>]><a/>',
      '<!DOCTYPE a [<!ATTLIST 1a b CDATA #IMPLIED>]><a/>',
      '<!DOCTYPE a [<!ATTLIST a 1b CDATA #IMPLIED>]><a/>',
      '<!DOCTYPE a [<!ATTLIST a b>]><a/>',
      '<!DOCTYPE a [<!ATTLIST a b FOO #IMPLIED>]><a/>',
      '<!DOCTYPE a [<!ATTLIST a b CDATA>]><a/>',
      '<!DOCTYPE a [<!ATTLIST a b CDATA x>]><a/>',
      '<!DOCTYPE a [<!ATTLIST a b CDATA "<">]><a/>',
      '<!DOCTYPE a [<!ATTLIST a b CDATA #FIXED"x">]><a/>',
      '<!DOCTYPE a [<!ATTLIST a b CDATA #IMPLIED"x">]><a/>',
      '<!DOCTYPE a [<!ATTLIST a b (x|) #IMPLIED>]><a/>',
      '<!DOCTYPE a [<!ATTLIST a b (x y) #IMPLIED>]><a/>',
      '<!DOCTYPE a [<!ATTLIST a b NOTATION x #IMPLIED>]><a/>',
      '<!DOCTYPE a [<!ATTLIST a b CDATA "&u;">]><a/>',
      '<!DOCTYPE a [<!ENTITY e SYSTEM "x"><!ATTLIST a b CDATA "&e;">]><a/>',
      '<!DOCTYPE a [<!ENTITY e "&#60;"><!ATTLIST a b CDATA "&e;">]><a/>',
      '<!DOCTYPE a [<!-- \u0001 -->]><a/>',
    ]) {
      expect(refusal(doc), doc).to.equal('XML/MALFORMED');
    }
    // A notation declaration split right after "<!NOTATIO" is still read whole.
    expect(refusal('<a><!NOTATION n SYSTEM "x"></a>', LIMITS, 12)).to.equal('XML/MALFORMED');
    expect(refusal('<!DOCTYPE a [<!ENTITY x "y">', LIMITS)).to.equal('XML/MALFORMED');
  });

  it('reads internal parameter entities, and skips the declarations after one it does not read unless the part is standalone', () => {
    const text = (doc: string) => events(doc)[1]?.v;
    // The parameter entity declares "t" first, so its declaration binds. Expat reads it so with parameter entity
    // parsing on, and libxml2 always.
    expect(text('<!DOCTYPE r [<!ENTITY % p "<!ENTITY t &#39;first&#39;>"> %p; <!ENTITY t "second">]><r>&t;</r>')).to.equal('first');
    expect(text('<!DOCTYPE r [<!ENTITY t "first"><!ENTITY % p "<!ENTITY t &#39;second&#39;>">%p;]><r>&t;</r>')).to.equal('first');
    // A reference between declarations in the text of a parameter entity, and comments and processing instructions.
    expect(text('<!DOCTYPE r [<!ENTITY % a "<!ENTITY t &#39;x&#39;>"><!ENTITY % b "<?p q?><!-- c --> &#37;a;"> %b;]><r>&t;</r>')).to.equal('x');
    // An external parameter entity is never fetched, and an undeclared one is not read either. The declarations
    // after it are skipped, so a reference to them is undeclared, unless the part is standalone.
    const unreadFirst = (ref: string, standalone: string) => `${standalone}<!DOCTYPE r [<!ENTITY % ext SYSTEM "none.ent">${ref}<!ENTITY t "after"><!ATTLIST r v CDATA "dflt">]><r>&t;</r>`;
    for (const ref of ['%ext;', '%undeclared;']) {
      expect(reason(unreadFirst(ref, '')), ref).to.equal('A reference to entity "t", which nothing declares');
      const doc = unreadFirst(ref, '<?xml version="1.0" standalone="yes"?>');
      expect(text(doc), ref).to.equal('after');
      expect(attributesOf(doc)[0], ref).to.deep.equal({ v: 'dflt' });
    }
    const { doctypes } = noted('<!DOCTYPE r [%ext;<!ENTITY t "after"><!ENTITY % p "x"><!ATTLIST r v CDATA "dflt"><!ELEMENT r ANY>]><r/>');
    expect(doctypes[0]).to.deep.include({ entities: 0, attributeDefaults: 0, skipped: 3 });
    expect(attributesOf('<?xml version="1.0" standalone="no"?><!DOCTYPE r [%ext;<!ATTLIST r v CDATA "dflt">]><r/>')[0]).to.deep.equal({});
    // A line end from a character reference in the text of a parameter entity stays as it is.
    expect(text('<!DOCTYPE r [<!ENTITY % p "<!ENTITY t &#39;&#13;&#10;&#39;><!ATTLIST r v CDATA &#39;&#13;&#10;x&#39;>">%p;]><r>&t;</r>')).to.equal('\r\n');
    expect(attributesOf('<!DOCTYPE r [<!ENTITY % p "<!ATTLIST r v CDATA &#39;&#13;&#10;x&#39;>">%p;]><r/>')[0]).to.deep.equal({ v: '  x' });
    expect(attributesOf('<!DOCTYPE r [<!ATTLIST r v CDATA "\r\nx">]><r/>')[0]).to.deep.equal({ v: ' x' });
    for (const [doc, why] of [
      ['<!DOCTYPE r [<!ENTITY % p "&#37;p;"> %p;]><r/>', 'Entity "%p" refers to itself'],
      ['<!DOCTYPE r [<!ENTITY % n "t"><!ENTITY % p "<!ENTITY &#37;n; &#39;x&#39;>"> %p;]><r/>', 'No white space after %'],
      ['<!DOCTYPE r [<!ENTITY % p "<!ENTITY t &#39;&#37;n;&#39;>"> %p;]><r/>', 'A parameter entity reference inside a declaration'],
      ['<!DOCTYPE r [<!ENTITY % p "<!ELEMENT r &#37;n;>"> %p;]><r/>', 'A parameter entity reference inside a declaration'],
      ['<!DOCTYPE r [<!ENTITY % p "<![INCLUDE[<!ENTITY t &#39;x&#39;>]]>"> %p;]><r/>', 'A conditional section, which only an external subset may hold'],
      ['<!DOCTYPE r [<!ENTITY % p "<!ENTITY t "> %p; &#39;x&#39;>]><r/>', 'An entity declaration with neither a value nor an external identifier'],
      ['<!DOCTYPE r [<!ENTITY % p "<!ELEMENT r ANY"> %p;]><r/>', 'A declaration in the internal subset does not end'],
      ['<!DOCTYPE r [<!ENTITY % p "<!NOTATION n SYSTEM &#39;x>"> %p;]><r/>', 'A declaration in the internal subset does not end'],
      ['<!DOCTYPE r [<!ENTITY % p "<!ENTITY t &#39;x>"> %p;]><r/>', 'An entity value is not closed'],
      ['<!DOCTYPE r [<!ENTITY % p "<!-- x"> %p;]><r/>', 'A comment in the internal subset does not end'],
      ['<!DOCTYPE r [<!ENTITY % p "<?p x"> %p;]><r/>', 'A processing instruction in the internal subset does not end'],
      ['<!DOCTYPE r [<!ENTITY % p "]"> %p;]><r/>', 'An unknown declaration in the internal subset'],
    ]) {
      expect(reason(doc), doc).to.equal(why);
    }
  });

  it('applies attribute defaults and normalizes attributes declared as tokens, as expat does', () => {
    const doc = (subset: string, body: string) => `<!DOCTYPE r [${subset}]>${body}`;
    expect(attributesOf(doc('<!ATTLIST r t CDATA "dflt" u CDATA #IMPLIED w CDATA #REQUIRED>', '<r/>'))[0]).to.deep.equal({ t: 'dflt' });
    expect(attributesOf(doc('<!ATTLIST r t CDATA "dflt">', '<r t="given"/>'))[0]).to.deep.equal({ t: 'given' });
    expect(attributesOf(doc('<!ATTLIST r t CDATA #FIXED "fixed">', '<r/>'))[0]).to.deep.equal({ t: 'fixed' });
    expect(attributesOf(doc('<!ENTITY e "v"><!ATTLIST r t CDATA "&e;x&#32;">', '<r/>'))[0]).to.deep.equal({ t: 'vx ' });
    expect(attributesOf(doc('<!ATTLIST r t CDATA "one"><!ATTLIST r t CDATA "two" u CDATA "more">', '<r/>'))[0]).to.deep.equal({ t: 'one', u: 'more' });
    expect(attributesOf(doc('<!ATTLIST r t (a|b) "a" u NOTATION ( n | m ) #IMPLIED v IDREFS " x  y ">', '<r/>'))[0]).to.deep.equal({ t: 'a', v: 'x y' });
    expect(attributesOf(doc('<!ATTLIST r t NMTOKENS #IMPLIED u CDATA #IMPLIED>', '<r t="  a   b  " u="  a   b  "/>'))[0]).to.deep.equal({ t: 'a b', u: '  a   b  ' });
    // A default can declare a namespace, as expat and libxml2 both apply it.
    expect(events(doc('<!ATTLIST p:r xmlns:p CDATA #FIXED "urn:p">', '<p:r/>'))[0].v).to.equal('{urn:p}r');
    // An element from entity text gets its defaults too.
    expect(attributesOf(doc('<!ENTITY e "<b/>"><!ATTLIST b t CDATA "dflt">', '<r>&e;</r>'))[1]).to.deep.equal({ t: 'dflt' });
    const { doctypes } = noted(doc('<!ATTLIST r t CDATA "dflt" u CDATA #IMPLIED>', '<r/>'));
    expect(doctypes[0]).to.deep.include({ attributeDefaults: 1, entities: 0 });
    expect(refusal(doc('<!ATTLIST r p:a CDATA "1">', '<r xmlns:p="urn:x" xmlns:q="urn:x" q:a="2"/>'))).to.equal('XML/MALFORMED');
  });

  it('limits entity nesting, expansions and expanded text, counted across the parts that share a budget', () => {
    expect(refusal(LAUGHS)).to.equal('LIMIT/XML_ENTITY_EXPANSIONS');
    expect(refusal(LAUGHS.replace('<lolz>&lol9;</lolz>', '<lolz v="&lol9;"/>'))).to.equal('LIMIT/XML_ENTITY_EXPANSIONS');
    // Quadratic blowup: one long entity referred to many times.
    const big = (body: string) => `<!DOCTYPE r [<!ENTITY big "${'x'.repeat(1000)}">]>${body}`;
    const size = { ...LIMITS, entitySize: 50_000 };
    expect(refusal(big(`<r v="${'&big;'.repeat(100)}"/>`), size)).to.equal('LIMIT/XML_ENTITY_SIZE');
    expect(refusal(big(`<r>${'&big;'.repeat(100)}</r>`), size)).to.equal('LIMIT/XML_ENTITY_SIZE');
    expect(refusal(big(`<r>${'&big;'.repeat(50)}</r>`), size)).to.equal('accepted');
    // A chain of references one deeper than the limit.
    expect(refusal(`<!DOCTYPE r [${chain(17)}]><r>&e1;</r>`)).to.equal('LIMIT/XML_ENTITY_DEPTH');
    expect(refusal(`<!DOCTYPE r [${chain(16)}]><r>&e1;</r>`)).to.equal('accepted');
    // Parameter entities count too.
    const pes = Array.from({ length: 17 }, (_, i) => `<!ENTITY % p${i + 1} "${i + 1 < 17 ? `&#37;p${i + 2};` : ''}">`).join('');
    expect(refusal(`<!DOCTYPE r [${pes}%p1;]><r/>`)).to.equal('LIMIT/XML_ENTITY_DEPTH');
    // Attribute defaults count as expansions of their text.
    const defaults = (n: number) => `<!DOCTYPE r [<!ATTLIST b t CDATA "${'x'.repeat(100)}" u CDATA #IMPLIED>]><r>${'<b u="1"/>'.repeat(n)}</r>`;
    expect(refusal(defaults(101), { ...LIMITS, entityExpansions: 100 })).to.equal('LIMIT/XML_ENTITY_EXPANSIONS');
    expect(refusal(defaults(100), { ...LIMITS, entityExpansions: 100 })).to.equal('accepted');
    expect(refusal(defaults(100), { ...LIMITS, entitySize: 9_999 })).to.equal('LIMIT/XML_ENTITY_SIZE');
    // Two parts that share a budget hit a limit that neither reaches alone.
    const part = `<!DOCTYPE r [<!ENTITY e "x">]><r>${'&e;'.repeat(60)}</r>`;
    const shared = { ...LIMITS, entityExpansions: 100 };
    expect(refusal(part, shared)).to.equal('accepted');
    const budget = { expansions: 0, size: 0 };
    expect(refusal(part, shared, 1 << 20, { budget })).to.equal('accepted');
    expect(refusal(part, shared, 1 << 20, { budget })).to.equal('LIMIT/XML_ENTITY_EXPANSIONS');
  });

  it('refuses encodings other than UTF-8 and UTF-16, and bytes that are not valid', () => {
    const cases: Array<[string, string | Uint8Array]> = [
      ['a Latin-1 declaration', '<?xml version="1.0" encoding="ISO-8859-1"?><a/>'],
      ['UTF-16 declared on UTF-8 bytes', '<?xml version="1.0" encoding="UTF-16"?><a/>'],
      ['invalid UTF-8', Buffer.from([0x3c, 0x61, 0x3e, 0xe9, 0x3c, 0x2f, 0x61, 0x3e])],
      ['UTF-32 big-endian', Buffer.from([0x00, 0x00, 0xfe, 0xff, 0, 0, 0, 0x3c])],
      ['UTF-32 little-endian with a byte order mark', Buffer.from([0xff, 0xfe, 0x00, 0x00, 0x3c, 0, 0, 0])],
      ['UTF-32 little-endian without one', Buffer.from([0x3c, 0, 0, 0, 0x61, 0, 0, 0])],
      ['UTF-8 declared on UTF-16 bytes', Buffer.concat([Buffer.from([0xff, 0xfe]), utf16le('<?xml version="1.0" encoding="UTF-8"?><a/>')])],
      ['UTF-16 with an odd byte at the end', Buffer.concat([Buffer.from([0xff, 0xfe]), utf16le('<a/>'), Buffer.from([0x3c])])],
    ];
    for (const [what, doc] of cases) expect(refusal(doc), what).to.equal('XML/ENCODING');
  });

  it('refuses what is not well-formed', () => {
    const bad = [
      '',
      '<a>',
      '<a></b>',
      '<a/></b>',
      '<a/><b/>',
      'text<a/>',
      '<a/>text',
      '<a/>&amp;',
      '<a b="1" b="2"/>',
      '<a xmlns:p="urn:x" xmlns:q="urn:x" p:b="1" q:b="2"/>',
      '<p:a/>',
      '<a p:b="1"/>',
      '<:a/>',
      '< a/>',
      '<a xmlns:p=""/>',
      '<a xmlns:xml="urn:other"/>',
      '<a xmlns:p="http://www.w3.org/XML/1998/namespace"/>',
      '<a xmlns="http://www.w3.org/XML/1998/namespace"/>',
      '<a xmlns:xmlns="urn:x"/>',
      '<a xmlns:1p="urn:x"/>',
      '<a b=1/>',
      '<a b/>',
      '<a b="1"c="2"/>',
      '<a b="<"/>',
      '<a <b/>',
      '<a / >',
      '<a>&amp</a>',
      '<a>&1;</a>',
      '<a>&#0;</a>',
      '<a>&#xD800;</a>',
      '<a>\u0001</a>',
      '<a b="\u0001"/>',
      '<a b="&amp"/>',
      '<a>&#xZ;</a>',
      '<a><!-- \u0001 --></a>',
      '<a:b:c xmlns:a="urn:a"/>',
      '<1a/>',
      ' <?xml version="1.0"?><a/>',
      '<?xml version="1.1"?><a/>',
      '<?xml encoding="UTF-8"?><a/>',
      '<a><?xml version="1.0"?></a>',
      '<?1x?><a/>',
      '<![CDATA[x]]><a/>',
      '<!FOO><a/>',
      '<![if x]><a/>',
      '<a',
      '<a><!-- x </a>',
    ];
    for (const doc of bad) expect(refusal(doc), JSON.stringify(doc)).to.equal('XML/MALFORMED');
  });

  it('refuses a reference left open past 64 characters while the text still streams in', () => {
    expect(refusal(`<a>&${'x'.repeat(100)}</a>`, LIMITS, 8)).to.equal('XML/MALFORMED');
  });

  it('limits element depth and token size, but not text', () => {
    const deep = `${'<a>'.repeat(6)}${'</a>'.repeat(6)}`;
    const small = { ...LIMITS, depth: 5, tokenSize: 1000 };
    expect(refusal(deep, small)).to.equal('LIMIT/XML_DEPTH');
    expect(refusal(deep, { ...small, depth: 6 })).to.equal('accepted');
    expect(refusal(`<a v="${'x'.repeat(2000)}"/>`, small)).to.equal('LIMIT/XML_TOKEN_SIZE');
    // A comment that arrives in pieces is refused once it outgrows the limit, before it ends.
    expect(refusal(`<a><!--${'x'.repeat(2000)}--></a>`, small, 100)).to.equal('LIMIT/XML_TOKEN_SIZE');
    expect(refusal(`<a>${'x'.repeat(5000)}</a>`, small)).to.equal('accepted');
    // A document type declaration is one token, however it arrives.
    expect(refusal(`<!DOCTYPE a [<!-- ${'x'.repeat(2000)} -->]><a/>`, small, 100)).to.equal('LIMIT/XML_TOKEN_SIZE');
  });

  it('reports the offsets of each tag in the decoded text', () => {
    const doc = '<r><a x="1">t</a><b/></r>';
    const seen: Array<[string, number, number]> = [];
    const t = new XmlTokenizer(
      {
        start: el => seen.push([`<${el.local}`, el.start, el.end]),
        end: (el, start, end) => seen.push([`/${el.local}`, start, end]),
      },
      { ...LIMITS, depth: 10, tokenSize: 100 },
      'x',
    );
    t.write(Buffer.from(doc));
    t.end();
    expect(seen.map(([name, start, end]) => `${name} ${doc.slice(start, end)}`)).to.deep.equal(['<r <r>', '<a <a x="1">', '/a </a>', '<b <b/>', '/b <b/>', '/r </r>']);
  });

  it("reads Office's VML dialect: both branches of a marker, and an unclosed <br> as an empty element", () => {
    const vml = '<xml><v:shape xmlns:v="urn:v"><![if gte mso 9]><v:imagedata/>\n<![endif]><div><b>One<br>Two<BR></b></div></v:shape></xml>';
    const names = (ev: Event[]) => ev.filter(e => e.e !== 'text').map(e => `${e.e} ${e.v.replace(/^\{[^}]*\}/, '')}`);
    const expected = ['start xml', 'start shape', 'start imagedata', 'end imagedata', 'start div', 'start b', 'start br', 'end br', 'start BR', 'end BR', 'end b', 'end div', 'end shape', 'end xml'];
    expect(names(events(vml, 1 << 20, LIMITS, { vml: true }))).to.deep.equal(expected);
    expect(names(events(vml, 3, LIMITS, { vml: true }))).to.deep.equal(expected);
    expect(refusal(vml)).to.equal('XML/MALFORMED');
    for (const doc of ['<xml><![foo]></xml>', '<xml><![if \u0001]></xml>', '<xml><![if x</xml>']) expect(refusal(doc, LIMITS, 1 << 20, { vml: true }), doc).to.equal('XML/MALFORMED');
  });

  it('tells XML names without a colon', () => {
    for (const s of ['a', '_a', 'a-b.c', '\xe9t\xe9', 'a\xb7', `x${String.fromCodePoint(0x300)}`, '\u{10000}x']) expect(isNcName(s), s).to.equal(true);
    for (const s of ['', '1a', '-a', 'a:b', '\xb7a', String.fromCodePoint(0x37e)]) expect(isNcName(s), s).to.equal(false);
  });
});
