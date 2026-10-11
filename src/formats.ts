export interface Format {
  /** The usual file extension, without the dot. */
  format: string;
  /** The MIME type of a file in this format. */
  mimeType: string;
}

const f = (format: string, mimeType: string): Format => ({ format, mimeType });

/** OOXML formats by the content type of the main part, lowercased. */
export const OOXML_FORMATS = new Map<string, Format>([
  ['application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml', f('docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')],
  ['application/vnd.ms-word.document.macroenabled.main+xml', f('docm', 'application/vnd.ms-word.document.macroEnabled.12')],
  ['application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml', f('dotx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.template')],
  ['application/vnd.ms-word.template.macroenabledtemplate.main+xml', f('dotm', 'application/vnd.ms-word.template.macroEnabled.12')],
  ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', f('xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')],
  ['application/vnd.ms-excel.sheet.macroenabled.main+xml', f('xlsm', 'application/vnd.ms-excel.sheet.macroEnabled.12')],
  ['application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml', f('xltx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.template')],
  ['application/vnd.ms-excel.template.macroenabled.main+xml', f('xltm', 'application/vnd.ms-excel.template.macroEnabled.12')],
  ['application/vnd.ms-excel.addin.macroenabled.main+xml', f('xlam', 'application/vnd.ms-excel.addin.macroEnabled.12')],
  ['application/vnd.ms-excel.sheet.binary.macroenabled.main', f('xlsb', 'application/vnd.ms-excel.sheet.binary.macroEnabled.12')],
  ['application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml', f('pptx', 'application/vnd.openxmlformats-officedocument.presentationml.presentation')],
  ['application/vnd.ms-powerpoint.presentation.macroenabled.main+xml', f('pptm', 'application/vnd.ms-powerpoint.presentation.macroEnabled.12')],
  ['application/vnd.openxmlformats-officedocument.presentationml.template.main+xml', f('potx', 'application/vnd.openxmlformats-officedocument.presentationml.template')],
  ['application/vnd.ms-powerpoint.template.macroenabled.main+xml', f('potm', 'application/vnd.ms-powerpoint.template.macroEnabled.12')],
  ['application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml', f('ppsx', 'application/vnd.openxmlformats-officedocument.presentationml.slideshow')],
  ['application/vnd.ms-powerpoint.slideshow.macroenabled.main+xml', f('ppsm', 'application/vnd.ms-powerpoint.slideshow.macroEnabled.12')],
  ['application/vnd.ms-powerpoint.addin.macroenabled.main+xml', f('ppam', 'application/vnd.ms-powerpoint.addin.macroEnabled.12')],
  ['application/vnd.openxmlformats-officedocument.presentationml.slide+xml', f('sldx', 'application/vnd.openxmlformats-officedocument.presentationml.slide')],
  ['application/vnd.ms-powerpoint.slide.macroenabled.main+xml', f('sldm', 'application/vnd.ms-powerpoint.slide.macroEnabled.12')],
  ['application/vnd.ms-visio.drawing.main+xml', f('vsdx', 'application/vnd.ms-visio.drawing')],
  ['application/vnd.ms-visio.drawing.macroenabled.main+xml', f('vsdm', 'application/vnd.ms-visio.drawing.macroEnabled.12')],
  ['application/vnd.ms-visio.template.main+xml', f('vstx', 'application/vnd.ms-visio.template')],
  ['application/vnd.ms-visio.template.macroenabled.main+xml', f('vstm', 'application/vnd.ms-visio.template.macroEnabled.12')],
  ['application/vnd.ms-visio.stencil.main+xml', f('vssx', 'application/vnd.ms-visio.stencil')],
  ['application/vnd.ms-visio.stencil.macroenabled.main+xml', f('vssm', 'application/vnd.ms-visio.stencil.macroEnabled.12')],
]);

/** OpenDocument formats by the media type in the mimetype entry or the manifest, exactly as written. */
export const ODF_FORMATS = new Map<string, Format>(
  (
    [
      ['text', 'odt'],
      ['text-template', 'ott'],
      ['spreadsheet', 'ods'],
      ['spreadsheet-template', 'ots'],
      ['presentation', 'odp'],
      ['presentation-template', 'otp'],
      ['graphics', 'odg'],
      ['graphics-template', 'otg'],
      ['formula', 'odf'],
      ['text-master', 'odm'],
      ['base', 'odb'],
    ] as const
  ).map(([kind, ext]) => [`application/vnd.oasis.opendocument.${kind}`, f(ext, `application/vnd.oasis.opendocument.${kind}`)]),
);
