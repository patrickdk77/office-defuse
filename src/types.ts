/** High-level kind of a finding. */
export enum OfficeCategory {
  Container = 'CONTAINER',
  Xml = 'XML',
  Limit = 'LIMIT',
  Package = 'PACKAGE',
  Encrypted = 'ENCRYPTED',
  Macro = 'MACRO',
  Ui = 'UI',
  Action = 'ACTION',
  Dde = 'DDE',
  Field = 'FIELD',
  External = 'EXTERNAL',
  Link = 'LINK',
  Embedded = 'EMBEDDED',
  Control = 'CONTROL',
  Processing = 'PROCESSING',
}

/** Exact variant of a finding. Some details appear under more than one category. */
export enum OfficeDetail {
  // CONTAINER
  NotZip = 'NOT_ZIP',
  AmbiguousEnd = 'AMBIGUOUS_END',
  MultiDisk = 'MULTI_DISK',
  DirectoryMismatch = 'DIRECTORY_MISMATCH',
  OutOfRange = 'OUT_OF_RANGE',
  Overlap = 'OVERLAP',
  HeaderMismatch = 'HEADER_MISMATCH',
  DuplicateName = 'DUPLICATE_NAME',
  UnsafeName = 'UNSAFE_NAME',
  NameEncoding = 'NAME_ENCODING',
  EncryptedEntry = 'ENCRYPTED_ENTRY',
  UnsupportedMethod = 'UNSUPPORTED_METHOD',
  Zip64Mismatch = 'ZIP64_MISMATCH',
  BadExtraField = 'BAD_EXTRA_FIELD',
  DataError = 'DATA_ERROR',
  LeadingBytes = 'LEADING_BYTES',
  HiddenData = 'HIDDEN_DATA',
  HiddenEntry = 'HIDDEN_ENTRY',
  TrailingBytes = 'TRAILING_BYTES',
  BadHeader = 'BAD_HEADER',
  BrokenChain = 'BROKEN_CHAIN',
  ChainLoop = 'CHAIN_LOOP',
  DirectoryLoop = 'DIRECTORY_LOOP',
  BadEntry = 'BAD_ENTRY',
  // XML
  Doctype = 'DOCTYPE',
  InternalSubset = 'INTERNAL_SUBSET',
  ExternalEntity = 'EXTERNAL_ENTITY',
  Malformed = 'MALFORMED',
  Encoding = 'ENCODING',
  // LIMIT
  FileSize = 'FILE_SIZE',
  EntryCount = 'ENTRY_COUNT',
  DecompressedSize = 'DECOMPRESSED_SIZE',
  TotalDecompressedSize = 'TOTAL_DECOMPRESSED_SIZE',
  XmlDepth = 'XML_DEPTH',
  XmlTokenSize = 'XML_TOKEN_SIZE',
  XmlEntityDepth = 'XML_ENTITY_DEPTH',
  XmlEntityExpansions = 'XML_ENTITY_EXPANSIONS',
  XmlEntitySize = 'XML_ENTITY_SIZE',
  RtfDepth = 'RTF_DEPTH',
  NestingDepth = 'NESTING_DEPTH',
  Time = 'TIME',
  // PACKAGE
  Unrecognized = 'UNRECOGNIZED',
  AmbiguousFormat = 'AMBIGUOUS_FORMAT',
  Unreadable = 'UNREADABLE',
  UnreadableBody = 'UNREADABLE_BODY',
  DuplicateEntry = 'DUPLICATE_ENTRY',
  UntypedPart = 'UNTYPED_PART',
  UnreferencedPart = 'UNREFERENCED_PART',
  OdfLayout = 'ODF_LAYOUT',
  // ENCRYPTED
  PasswordRequired = 'PASSWORD_REQUIRED',
  DefaultPassword = 'DEFAULT_PASSWORD',
  UserPassword = 'USER_PASSWORD',
  Rc4_40 = 'RC4_40',
  Rc4_128 = 'RC4_128',
  Rc4Other = 'RC4_OTHER',
  Aes128 = 'AES_128',
  Aes192 = 'AES_192',
  Aes256 = 'AES_256',
  UnsupportedAlgorithm = 'UNSUPPORTED_ALGORITHM',
  XorObfuscation = 'XOR_OBFUSCATION',
  // MACRO
  VbaProject = 'VBA_PROJECT',
  CompiledCode = 'COMPILED_CODE',
  XlmMacrosheet = 'XLM_MACROSHEET',
  WordBasic = 'WORDBASIC',
  OdfBasic = 'ODF_BASIC',
  OdfScripts = 'ODF_SCRIPTS',
  OdfDialogs = 'ODF_DIALOGS',
  EventOnOpen = 'EVENT_ON_OPEN',
  Event = 'EVENT',
  Link = 'LINK',
  // UI
  Ribbon = 'RIBBON',
  Callback = 'CALLBACK',
  OdfConfiguration = 'ODF_CONFIGURATION',
  // ACTION
  RunProgram = 'RUN_PROGRAM',
  SubmitForm = 'SUBMIT_FORM',
  OleVerb = 'OLE_VERB',
  ObjectUpdate = 'OBJECT_UPDATE',
  // DDE
  Field = 'FIELD',
  Formula = 'FORMULA',
  // FIELD
  Macrobutton = 'MACROBUTTON',
  Computed = 'COMPUTED',
  // EXTERNAL
  Template = 'TEMPLATE',
  Object = 'OBJECT',
  Image = 'IMAGE',
  Frame = 'FRAME',
  Include = 'INCLUDE',
  Workbook = 'WORKBOOK',
  Reload = 'RELOAD',
  Other = 'OTHER',
  DataConnection = 'DATA_CONNECTION',
  LocalFile = 'LOCAL_FILE',
  // LINK
  Safe = 'SAFE',
  FileUrl = 'FILE_URL',
  NetworkPath = 'NETWORK_PATH',
  DataUrl = 'DATA_URL',
  ScriptUrl = 'SCRIPT_URL',
  OtherScheme = 'OTHER_SCHEME',
  Credentials = 'CREDENTIALS',
  IpHost = 'IP_HOST',
  LookalikeHost = 'LOOKALIKE_HOST',
  EncodedHost = 'ENCODED_HOST',
  Relative = 'RELATIVE',
  TextMismatch = 'TEXT_MISMATCH',
  // EMBEDDED
  OleObject = 'OLE_OBJECT',
  Package = 'PACKAGE',
  OleFile = 'OLE_FILE',
  Applet = 'APPLET',
  Plugin = 'PLUGIN',
  Equation = 'EQUATION',
  Attachment = 'ATTACHMENT',
  PluginPassed = 'PLUGIN_PASSED',
  PluginScrubbed = 'PLUGIN_SCRUBBED',
  PluginFailed = 'PLUGIN_FAILED',
  PluginRemoved = 'PLUGIN_REMOVED',
  // CONTROL
  ActiveX = 'ACTIVEX',
  FormsControl = 'FORMS_CONTROL',
  ControlPicture = 'CONTROL_PICTURE',
  // PROCESSING
  VerificationFailed = 'VERIFICATION_FAILED',
}

export type OfficeFindingAction = 'reject' | 'strip' | 'info';

/**
 * A finding as any package of the defuse family reports it. Category and detail are plain strings, since the string
 * enums of two packages never match each other.
 */
export interface DefuseFinding {
  category: string;
  detail: string;
  /** Plain sentence, safe to show the uploader. */
  description: string;
  /** Default action after any caller override. */
  action: 'reject' | 'strip' | 'info';
  /** Where in the file, such as a part name or a relationship. */
  location?: string;
  /** Extra values: the target, the field code, the content type, counts. */
  data?: Record<string, string | number>;
  /**
   * Set when the finding comes from inside a contained file: the file's name, with " > " between nesting levels.
   * `location` then starts with where that file sits.
   */
  attachment?: string;
  /** How much the finding adds to a score, set by the package that made it. */
  weight?: number;
}

/** A finding office-defuse makes itself. */
export interface OfficeFinding extends DefuseFinding {
  category: OfficeCategory;
  detail: OfficeDetail;
}

export enum OfficeRisk {
  None = 'NONE',
  Low = 'LOW',
  Medium = 'MEDIUM',
  High = 'HIGH',
  Critical = 'CRITICAL',
  /** The file could not be inspected: a refusal of the container or XML layer, or a limit, stopped the run. */
  Unknown = 'UNKNOWN',
}

export interface OfficeInspection {
  status: 'clean' | 'strippable' | 'rejected';
  /** 0 to 100, null when risk is UNKNOWN. */
  score: number | null;
  risk: OfficeRisk;
  findings: DefuseFinding[];
  /** The format the content says, as its usual file extension without the dot, such as 'docx' or 'ods'. '' when not recognized. */
  format: string;
  /** The MIME type of a file in that format. '' when not recognized. */
  mimeType: string;
}

export interface OfficeDisarmResult {
  status: 'clean' | 'defused' | 'rejected';
  /** The uploaded file and its score. */
  before: OfficeInspection;
  /** The output and its remaining score. Absent when rejected. */
  after?: OfficeInspection;
  /** Findings whose content was removed. */
  removed: DefuseFinding[];
}

/** Random-access input. A buffer, a local file and an S3 ranged GetObject all fit. */
export interface ByteSource {
  size(): Promise<number>;
  read(offset: number, length: number): Promise<Uint8Array>;
  close?(): Promise<void>;
}

/** Append-only output. Never seeks. */
export interface ByteSink {
  write(chunk: Uint8Array): Promise<void>;
  close(): Promise<void>;
  /** Called instead of close() when copying to the sink fails, so a partial output is not taken for a whole one. */
  abort?(error: unknown): Promise<void>;
}

export interface OfficeLimits {
  /** Bytes in the upload. Not enforced unless set. */
  fileSize?: number;
  /** Entries in the ZIP central directory. Default 10,000. */
  entries?: number;
  /** Declared decompressed size of any one entry, in bytes. Default 512 MiB. */
  decompressedBytes?: number;
  /** Declared decompressed sizes of all entries together, in bytes. Default 2 GiB. */
  totalDecompressedBytes?: number;
  /** Element nesting depth in one XML part. Default 256. */
  xmlDepth?: number;
  /** Characters in one XML tag, comment, processing instruction or CDATA section. Default 16 Mi. */
  xmlTokenSize?: number;
  /** Entity references nested inside entity text. Default 16. */
  xmlEntityDepth?: number;
  /** Entity references expanded in the whole run. Default 100,000. */
  xmlEntityExpansions?: number;
  /** Characters of entity text expanded in the whole run. Default 16 Mi. */
  xmlEntitySize?: number;
  /** Group nesting in an RTF file, and destinations open at once. Default 256. */
  rtfDepth?: number;
  /**
   * The deepest a file may sit in the upload. The upload is at depth 0 and a file inside it at 1, and files of every
   * format in between count. Not enforced unless set.
   */
  nestingDepth?: number;
  /** Milliseconds for the whole run. Not enforced unless set. */
  timeMs?: number;
}

export interface FormatDetection {
  /** The format's MIME type, such as application/vnd.openxmlformats-officedocument.wordprocessingml.document. */
  mimeType: string;
  /** The format's usual extension, lowercase and without the dot. */
  extension: string;
  /** Where the format starts in the file: 0, or the offset of its data when other data comes first. */
  offset: number;
  /** Set when the match rests on more than the content, and says what it rests on. */
  heuristic?: string;
}

/** What a detect function may know of a file besides its content. */
export interface DetectOptions {
  /** The file's name, as the upload or the container that holds it gives it. */
  name?: string;
  /** The MIME type the upload or the container declares. */
  declaredType?: string;
  /** The Date.now() value detection must finish by, as `PluginContext.deadline`. */
  deadline?: number;
  /** Folder for temporary files. Defaults to `os.tmpdir()`. */
  tempDir?: string;
  /** Bytes held in memory before a temporary file takes over. */
  memoryThreshold?: number;
}

/** What a contained file plugin learns of the run, with nothing format-specific, so any package of the family can make one. */
export interface PluginContext {
  /** The depth of the file itself: 0 for the upload, 1 for a file inside it, and so on, whatever the formats. */
  depth: number;
  /** When a time limit is set, the Date.now() value the whole run must finish by. Nested runs share it. */
  deadline?: number;
}

export interface ContainedFile {
  name?: string;
  /** The MIME type the container declares for the file, such as a part's content type. */
  declaredType?: string;
  /** From the first bytes. */
  sniffedType?: string;
  /** Decoded size. */
  size: number;
  location: string;
  /** The depth of the file itself: 0 for the upload, 1 for a file inside it, and so on, whatever the formats. */
  depth: number;
  /** Decoded bytes. */
  source: ByteSource;
}

/** What a file plugin did, plus anything it found inside the file. */
export interface ContainedFileResult {
  result: 'scrubbed' | 'passed' | 'removed';
  /** Findings inside the file as it arrived. They go into the report of the file that contains it. */
  findings?: DefuseFinding[];
  /** Findings inside the bytes the plugin kept. They go into the report of the output. */
  outputFindings?: DefuseFinding[];
}

export interface ContainedFilePlugin {
  kind: 'file';
  name: string;
  accepts(file: ContainedFile): boolean | Promise<boolean>;
  /** Writes the bytes to keep. Returning 'removed' drops the file after all. */
  process(file: ContainedFile, sink: ByteSink, context: PluginContext): Promise<'scrubbed' | 'passed' | 'removed' | ContainedFileResult>;
}

/** A type library or project a VBA project refers to, as its dir stream lists it. */
export interface VbaReferenceInfo {
  name: string;
  kind: 'registered' | 'control' | 'project';
  /** A type library's LibidReference, such as `*\G{...}#2.0#0#path#description`, or a project's path. */
  libid: string;
}

/** A VBA module, as office-defuse shows it to a script plugin. */
export interface ContainedScript {
  /** The module's source, decoded with the project's code page. */
  text: string;
  /** "vba-" and the module's kind: "vba-document", "vba-standard", "vba-class" or "vba-form". */
  trigger: string;
  location: string;
  /** The module's name. */
  fieldName: string;
  language: 'vba';
  form: 'office';
  /** The references of the module's project. office-defuse gives them with the project's first module only. */
  references?: VbaReferenceInfo[];
}

export interface ScriptPlugin {
  kind: 'script';
  name: string;
  accepts(script: ContainedScript): boolean | Promise<boolean>;
  /**
   * 'passed' keeps the module. Anything else removes the VBA project, since office-defuse writes no module a plugin
   * rewrote. `findings` go into the report at the module's location.
   */
  process(script: ContainedScript): Promise<{ result: 'scrubbed' | 'passed' | 'removed'; text?: string; findings?: DefuseFinding[] }>;
}

export interface OfficeActionOverride {
  category: OfficeCategory;
  detail?: OfficeDetail;
  action: OfficeFindingAction;
}

export interface OfficeScoreWeight {
  category: OfficeCategory;
  detail?: OfficeDetail;
  weight: number;
}

/** Lower bound of each band. */
export interface OfficeScoreBands {
  low: number;
  medium: number;
  high: number;
  critical: number;
}

export interface OfficeOptions {
  /**
   * One password or a list, tried in order on an encrypted file: an encrypted OOXML package, or an .xls workbook.
   * VelvetSweatshop, the password Excel opens a workbook with by default, comes after them. Pass passwords in code,
   * never on a command line, where other users can read them.
   */
  password?: string | string[];
  limits?: OfficeLimits;
  actionOverrides?: OfficeActionOverride[];
  scoreWeights?: OfficeScoreWeight[];
  scoreBands?: OfficeScoreBands;
  /** Plugins that decide about the files inside the file: embedded objects and packages, OLE parts and attachments. */
  filePlugins?: ContainedFilePlugin[];
  /**
   * Plugins that decide about each module of a VBA project in a compound file or an OOXML package. The first plugin that
   * accepts a module decides it. The project stays when every module passed, written without its compiled code so
   * Office compiles the source the plugins read. Otherwise it goes, as it does without plugins.
   */
  scriptPlugins?: ScriptPlugin[];
  /** Directory for temporary files. Defaults to os.tmpdir(). */
  tempDir?: string;
  /**
   * Default 8 MiB. disarmOfficeSource keeps its copy of the upload and its output in memory up to this size and in a
   * temporary file above it. Each file inside that goes to a plugin does the same, and the copies plugins return,
   * which wait for the output to be written, share this much memory.
   */
  memoryThreshold?: number;
}
