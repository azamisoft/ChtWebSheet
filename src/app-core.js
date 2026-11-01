// Shared bindings and ordered application initialization.
import {
  registerWebSheetFormulaPlugin,
} from "./modules/agent/contracts.js";
import {
  formulaAutocompleteShouldExposeEngineFunction,
  formulaDefaultTemplate,
  formulaEngineFunctionCategory,
  formulaEngineFunctionDescription,
  formulaEngineFunctionMetadata,
  formulaEngineFunctionSyntax,
} from "./modules/formula/autocomplete.js";
import {
  warmWorkbookIoModules,
  webSheetFormulaError,
} from "./modules/formula/calculations.js";
import {
  formulaTemplateReference,
  recentFormulaFunctionNames,
} from "./modules/formula/dialogs-and-names.js";
import {
  createWebSheetFormulaPluginClass,
} from "./modules/formula/plugin.js";
import {
  paintSheetCanvas,
  renderCanvasLayer,
} from "./modules/rendering/grid-and-canvas.js";
import {
  firstCellReferenceFromRangeLabel,
} from "./modules/ui/ribbon-and-links.js";
import {
  updateWorkbookNameDisplay,
} from "./modules/ui/views-and-formula-tools.js";
import {
  hyperFormulaBuildConfig,
} from "./modules/workbook/formulas-and-sheet-tabs.js";
import {
  init,
} from "./modules/workbook/startup-and-loading.js";
import $ from "jquery";


import * as XLSXModule from "@e965/xlsx";


import ExcelJSModule from "exceljs";


import { ArraySize, CellError, ErrorType, FunctionArgumentType, FunctionPlugin, HyperFormula, SimpleRangeValue } from "hyperformula";


import JSZipModule from "jszip";


import {
  AlignCenter,
  AlignHorizontalJustifyCenter,
  AlignHorizontalJustifyEnd,
  AlignHorizontalJustifyStart,
  AlignJustify,
  AlignLeft,
  AlignRight,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  ArrowDownNarrowWide,
  ArrowDownToLine,
  ArrowUpDown,
  BadgePlus,
  BarChart3,
  BookOpen,
  Box,
  CalendarClock,
  CaseSensitive,
  ChartColumn,
  ChartColumnBig,
  ChartLine,
  ChartNoAxesColumn,
  ChartPie,
  Check,
  ChevronDown,
  ChevronUp,
  ChevronsDown,
  ChevronsUp,
  CircleHelp,
  ClipboardPaste,
  Clock3,
  Columns2,
  Columns3,
  Copy,
  Crop,
  Database,
  Delete,
  Download,
  Eraser,
  Eye,
  EyeOff,
  FilePlus,
  FileSpreadsheet,
  FileText,
  FileUp,
  Filter,
  FolderOpen,
  FunctionSquare,
  GitBranch,
  Globe,
  Grid2X2,
  Hash,
  Image as ImageIcon,
  Info,
  Italic,
  KeyRound,
  Link as LinkIcon,
  LogOut,
  ListChecks,
  ListFilter,
  MailCheck,
  Map as MapIcon,
  Maximize,
  MessageSquareText,
  MessageSquareX,
  MousePointer2,
  Minus,
  MonitorUp,
  MoveRight,
  Omega,
  Paintbrush,
  PaintBucket,
  Palette,
  PanelTop,
  PenLine,
  Pencil,
  Highlighter,
  Lasso,
  Plug,
  Printer,
  Puzzle,
  RefreshCw,
  Repeat,
  RotateCcw,
  RotateCw,
  Ruler,
  Save,
  SaveAll,
  Scissors,
  ScatterChart,
  Search,
  Send,
  Settings,
  Shapes,
  Signature,
  Sigma,
  Square,
  SquareDashed,
  SquareFunction,
  Sparkles,
  Table2,
  TableCellsMerge,
  TableCellsSplit,
  TableProperties,
  TextCursorInput,
  Trash2,
  Type,
  Underline,
  UnlockKeyhole,
  Unlink,
  Upload,
  UploadCloud,
  WrapText,
  X,
  PlaySquare,
  ZoomIn,
  createIcons,
} from "lucide";


import { initCwsAgentBridge } from "./cws-agent-bridge.js";


import "./styles.css";



const ROW_HEADER_MIN_WIDTH = 28;


const ROW_HEADER_MAX_WIDTH = 72;


const ROW_HEADER_DIGIT_WIDTH = 7;


const ROW_HEADER_PADDING_WIDTH = 14;


const COLUMN_HEADER_HEIGHT = 22;


const FALLBACK_VISIBLE_ROWS = 40;


const FALLBACK_VISIBLE_COLS = 20;


const SCROLL_EXTEND_ROWS = 80;


const SCROLL_EXTEND_COLS = 8;


const SCROLL_EXTEND_THRESHOLD = 120;


const SCROLL_TRIM_BUFFER_PX = 600;


const SCROLL_TRIM_IDLE_MS = 180;


const IMAGE_RENDER_BUFFER_PX = 160;


const TEXT_OVERFLOW_LOOKBACK_MIN_PX = 2400;


const TEXT_OVERFLOW_LOOKBACK_MAX_PX = 8192;


const TEXT_OVERFLOW_LOOKBACK_VIEWPORT_FACTOR = 2;


const SELECTION_AUTOSCROLL_EDGE_PX = 34;


const SELECTION_AUTOSCROLL_MIN_STEP_PX = 2;


const SELECTION_AUTOSCROLL_MAX_STEP_PX = 26;


const VIRTUAL_ROW_OVERSCAN = 8;


const VIRTUAL_COL_OVERSCAN = 4;


const FORMULA_BAR_MIN_LINES = 1;


const FORMULA_BAR_DEFAULT_EXPANDED_LINES = 3;


const FORMULA_BAR_MAX_LINES = 12;


const ANNOTATION_POINTER_TIP_OFFSET = 20;


const ANNOTATION_POINTER_HALF_HEIGHT = 8;


const EMU_PER_PIXEL = 9525;


const LINE_SHAPE_MIN_BOX_SIZE = 12;


const LINE_SHAPE_BASE_PADDING_PX = 4;


const LINE_SHAPE_ARROW_PADDING_EXTRA_PX = 3;


const LINE_CONNECTION_RETENTION_TOLERANCE_PX = 4;


const HYPERFORMULA_LICENSE = "gpl-v3";


const VERSION_INFO_FALLBACK_VERSION = "0.0.0";


const STANDALONE_RUNTIME_VERSION = VERSION_INFO_FALLBACK_VERSION;


const STANDALONE_RUNTIME_CHANNEL = "current";


const STANDALONE_IMAGE_ASSET_ENCODING_GZIP_BASE64 = "gzip-base64";


const STANDALONE_IMAGE_ASSET_GZIP_MIN_BASE64_LENGTH = 8192;


const STANDALONE_IMAGE_ASSET_GZIP_MIN_SAVINGS_RATIO = 0.98;


const STANDALONE_IMAGE_MAX_DIMENSION = 3000;


const STANDALONE_IMAGE_DEVICE_PIXEL_RATIO = 2;


const STANDALONE_IMAGE_WEBP_QUALITY = 0.88;


const APP_DISPLAY_NAME = "Cht WebSheet";


const APP_SHORT_NAME = "CWS";


const APP_COMPANY_NAME = "株式会社CHT";


const APP_VERSION = VERSION_INFO_FALLBACK_VERSION;


const APP_ICON_VERSION = "20260612-margin1";


const APP_ICON_DATA_URL = [
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAATpElEQVR42u1beZgU1bX/3bpVvU",
  "z1TPfse88MDDsoiLIODCBIZDEuJMYlanxqQhJRccvTR9S4RCUxT6PRxHwxiWgQAYkx+kRFcQOJKKAogwiKCgPD7NNLdVXd8/",
  "7oqrJ6pmeYnqhviff76uvpnrq37u/s59xTDF/RICIJALOuPm8FQIwx8VXsi32JgBkACQBjjBkDXEO2CCIYY/R/ggBExJFEbX",
  "b7vQxADYBqAKUA8gD4rX/HALQAOAjgIwD7GGMH+rPu/xoCWBt0OEVE+QDqAcwBMAlALYBAP5frArAHwGYAzwHYyBhrdkvWF0",
  "2If0q3rU3Z32cS0R+J6BD1HIKIDCLSe7kM657u45C15ky3ill25X8MOLPF0vp+KhG93G3jNlizF2C9DWHNsYniHi8T0aluyX",
  "Mz4Ku06PbfE4joOdcGzT44OdBhS47p+u05IpqQbk9fNnhufSpEdKvFJZvbBn35w/0c3dqD4t7blwletj5riei1btzpm4VCpF",
  "ymafa4ut/TD0LYN71GRLXuPX6Z4GcR0WEXB3oANU2TdMMgXdfJNM1/it3d10pDGHsPh4loVqZEYP0FzxgziGgRgEcBKABMAI",
  "7ICSEgiCDz9FKoaQlEIhFoCQ0kCLF4HEQEZkU6jDH4fT6AAR7FA3+WH2pWVtq1TNMEYwyS5Ki9vRcdwNmMsdX2nv9pArjAnw",
  "HgcVe4KtnAATibEULg/V0N2LZtB97duRMf7/8EBw82oq29HR0dHUgkdBARdD356R4ejwcMgKzIULNU5ObloqS4CJWVFRg9ai",
  "SOGzcWo0aOgKIozrMIAE8+W7jwfIsxtqY/RGBHM3iMMdPyv+tt0AAkIoIQBM6TP+14512sXrsOzz+/AQ27P0BnZyeICBLnkL",
  "kMLnNwiUOSWArBrOgORJRymULANAwYhgnTTMY8gYCKwYNqMGf2ifjWotNw3LixDiEYY3DlDwLASYyxF20MGROAiCTGmCCiGg",
  "BvWqGrACCZpgluifqmzVtw9z33Yf3zL6CjsxN+nw8+nw+ccwcYs0DC+i6EgGkKCGE6gBmTIEkMnHNIkgRJkhz1SEoKwTQF4v",
  "E4YrEYVFXFrJn1uHzJjzB9Wp2jGpxzAUASQrRIknQ8Y2yfjaXfBHAlMhKAVwFMsPXMBn+4qQk/vfFmPPKXVdATCeTk5IBznr",
  "QFFkc4lyAEIR7XoGlxGEZyrsfrgZqVBb/fB1lW4PV4oCUSSCQ0dHVFEYvFoOsJMMbg9Xrh8/kgy3KKupmmiY7OTsicY9EZp+",
  "GWm36Kiopy2z6YkiRxAFsA1FmMS5tQsaOI/q0ArgNgAJBt8M+ufx6XLb0aez7ci/y8PDDGHDGVJAmMMcTjcUQiUXi9XlRXh3",
  "HMmNE49pgxGFI7GBXl5SgoLEB2dgCKosCjKEgkdGiahuaWFjQ2HsKHe/dix453sfXtbWho2I2Ojg74fD6oqupIEeccRISW1l",
  "ZUlJXhl7+4HaefeopNKEOSJBnAbYyx63tTBdYH+HEA/pHUMUiAYJIkYfkv/xM33HQLFEVBVlYWDMNw9JhLEroiEWiahtrawV",
  "i4YB4WzPsGxo0di0BAHZD7FULgvfd34Zn/Wo81T/wV27bvAJc4cnICME0BIoIsy4jH44hGY7j26ivwsxuXARAkBIRla05gjL",
  "19NHuQElIS0UuWHzZsX37FVdcS8wSoqKyaisqqqaAkTAUlYSqpGEShgjJS/EGaOHUGPfTHh6mtra2nP9d1MgyDDCN94GPHEI",
  "bl9w0jNb5K6Do9se5Jqp81l2R/kHILy6mkooYKSsJUVFZNxeU1xJQAXXjxYnu+vcBL/QqXXWHufBu8vcblV15DUFQqqRhEha",
  "VVVFASpsLSKiouryHFH6KaIaPogd/9njRN+zxUs8Cmi+oMiyC6rvcZ9QlXMOQm5oO/f4iqB48gT1YulVQMcphRWjmYwLPo+z",
  "9cYu/Bnjj/qOGyndoS0SYiErquG0REt92+3AFvP6i4vIbyiytJ8Yfo/Asvoc8+O5ACvE9QQgw8IzIMsqfv3/8JLTrzHOK+HC",
  "oqq3YYU1o5mCCr9Ot773eHzJv6TKFd3J9GRKTruklE9LennibFH6Li8hrnASUVgygQKqK8ogp68PcPfR6T6jodDZutThte3E",
  "i33/lLWn7X3fTJp59Z4n90wtjqEY/Hnd/uWH4X+bPzKb+4MoUQ+cWV1LD7A3JlkdN6lQIXAf5sUUxvOnKEaoeNoWB+qaPzJR",
  "WDyKvm0THjJtLO997PiLO2nkeiURo+ehyB+wkAXb70Ggfc0eb3Nh5btZqC+aWUV1RBRaVVVFIxiGR/kC5ZfKmdLwgi+nN3Aj",
  "Db7zPGiIhCAD60gh66btmN7I7ld6G4qAiJRAIeRcGR5hbUTZ2MlY/8CZFIBAcbG6EoCgoKClBTXdWnjTFMEzLnWPHoSlx48W",
  "JnXcYYNr/6IqqqwsnoUZLSMQiMMbz2+ia8vukNHDh4EFpcQzhciauWXgZZlvH0M8/iO+dcAK/XA0mSkEgkkJ+fhzc3v0KhUI",
  "hZwdFgxlibjbl7pneaLWmHm5qoavAIyi0sp0IXReefcgZ1dXUREdGM2SeT5M2m3MJyChWU0aWXXelY8t44mEjoNLluFqmhIi",
  "oqq06u6wvSjy+7slcpsH/79X0PEBSVFH+QPGouAQqNO35y0kMkEkRE9Mc/rSCPmkvl4VoqLK2imiGj6ODBRnKl7Kf1yBhdBH",
  "jAFv89e/dRWbiWistrqKJqCHnUXJq38HSKRmMOmLnzvklqsNAxjswToBc2vJQWiP19zdq/OjbFNqi2vu5q2N2DgLbYd3V10b",
  "CRYymYX0pl4cFUWjmYgvmlVFc/+3M3aRHh/O9dTAAnxR+kYEGZTQBbDR5wY7apYFrh7yRLLSQQobm5BaZpwtDiOGHiRPxlxU",
  "Pw+33QdR2KoiAvLxeGYVjBCIfP68U9996PWTPrk7G/K3W0o8Vf3f1reDxKSiYoyzKam1twx/K78IcH73dCXjsQ4pzjpY2v4K",
  "OP9yMUCiKR0ME5RzQaxejRo9x5AIgIN9/0U3h9PkgSg6qqCAQCsM8oAEyysJrdA59SIorYhG9tbaVlN9xM1/3HDXTVNdfRnj",
  "0funx7kpvrnnyKPFm5DjeLyqopO7eYtvxjawrX7c9Vj68lxR+ikm7ct2OK7LwS2vrW22nnXnjRD0jJ+lxyCkurKCe/lLZt39",
  "EvA+qqHkWIqNTB7rL+U7vd2KsVtj87Oztp+OhxFMwvTdqJ8hpSskJ03vcutjZlOiKtaRqdMHk6BSzd706A4vIa8qp5dNqis3",
  "rEEi2trTR42GgKFZRRYWkVFZVVkz+ngGbPXZDWOwghSLcY1Y0w9o1TbW8gufKBald1BUSAYRjOZWd4cIlzIBDASbNPRCQSBZ",
  "ckGKaJYE4Onvr7M2jY/QE4l6DrBiRJwiOPPoa33t6GQCCQIuLuKk8oFMTTzzyL9c+9AM45ElbR5LnnXsD+/Z/C6/U6XiKhJf",
  "CNuXOcuSkJDmOQOQe3LvdjumFlbn9TlrpIUjftK61rArBgwcmQZRkEcvS5o6MTv7n/dwAAziW0t7fjjl/chYCqpmzWKmKkuD",
  "pZ5rj51tthGAa4lVmueeKvTiHFBqyqWZg1c0aP4ko/h4PVPTM3kxXsgsXkiRNQU1OFeFxzJCMYzMGq1Wuxb99H4JzjV3ffhz",
  "179sLv9zvGjzEGTdOg67pDBCEEsrOzsemNLXh4xV8gyzL2fLgXG19+1ZIcE5IkIR6PY+jQIRg1cvhACZCbjgD+jA4VXWowo3",
  "46otEoJEkCEUFRFBw6fBgrV63GoUOHcc+9v0EoFHS4zzlHe3sHFp2RLGt1dUUcEKZpIqCquOXnd0LXdTy+ei2ampqgKAqIks",
  "+NxeKYPGkiPB5PD/Hv5/CnI8CAx/yT56ZwwTRN5GRn4+FHVuK7F1yERCLhuCj7/z6fD7fefAOWXrEEWkJz5hMRfD4fDhw4gC",
  "WXX4XVa9ZB7aY6ksQwo37aF1LudxMglvFka9N1U6cgHK6EpmlOHdDj8aCxsRGb39gCVVUdwyfLMtra23H2Wd9GSXEx5s45Ef",
  "XT6tDR0ZEiBTk5OXj4kZXY99FH8Pl8Tiis6zqKi4tRN2XSQMU/Bat7dmvGZ+sunZ8+rS7pDbjkMmgy/H5/itU3DAPBYBBLL7",
  "/Uuef6f78GpsvL2PZAVbOSBtaSHEmSEIlEMWniBBQVFaV4pgxHazoCHBjQiZEFduH8ky3upyYwbvCyLKOtrR3nnPVtDKkd7F",
  "SHT5w1A6csmI+2trYUtyWESIkYbYKfsnAe3GcSAxgOVuaqAU61KsCUSeOELZpt7e2YMGk6mo4csQwW9ZAWIQQURcGW1zeioq",
  "I8hbPvvb8LU6fPhiJzUC/SZhgGQsEgtmx+GQX5+c6zM+QXA1DHGHvNDoTs5+0FELVuoEzVIBQMoq5uCiKWN+g+OOdobWvD4h",
  "9cjMrKCgghnPq/aZoYOWI4Fn//IjS3tkKR5bTzOzs7sXDhPBTk5zvHYwMAH7WwAoBz+MEANAL4wHVzBgenyc9vLpwPlkZ4JE",
  "lCNBrDsKFDcPmlP3TAu/8vhMBPrr0SI4cPRyQahSzLTqBkcz8YDOJHP7jEIfwAtBUWxkarHiDsXXCrOLDZ7srKZGXb8M2cUZ",
  "8SFKUSIIqf3bgMOTk5PUTX9hw52dn4xfLbEI3GEI1GYZpmkliM4ciRZlxw/rkYOnQITNMciPUXFrbNFlaeLg54Nl0vn23M3F",
  "c6NcjJycZJc05EJBJxiKIoCo4cOYJzzj4Ti04/NSVtFUI4uQYAJBIJzJ0zG8uuvxYFBQXw+/3JyC+RwJgxo/CTa67sIT0ZNo",
  "QxC2PaozAQUYiImo+WFabLwOys65VXXyM1WEhl4Voqr6olNVhEY8ZOoOaWlmTN36r792dEIhFqb2+nlpYW0nXdea6d5WVQXb",
  "ZvbLbKfg5mluZE6M8AzrWKJDJjDI2Nh7DkiqshhIlYLIZl1/8EkyZOgPuQ1JYUUwhMnTYLWzZtguT1oSA/H+ufeRLHHjPGKa",
  "TYwc72He/grbe3Y9++fWg8dBiGYYCBQVCSy6FQEEWFhQiHK1FdVYVwuBLhyooeWaR9HNdXOdIS+RWMsfPcJ0TpCDANwMtJNy",
  "skSZKwa1cDxp4wBRJjiMc1jBwxHC+/tB55ubkwDBOyzB2/LEkSNry4ES9seBGMSThlwTxMmHC8A769vQMP/elhrFy1Gu+/vw",
  "vRaAwEcpIrx1qR3XQhwMDg9XqQm5uLoUNqMWXKJMyeNRNTp0yGx6P0hxDCUvfpjLFX0hLAVR0iAK8DmGiapuCc8127GlB/4j",
  "ecPLu5pQX106dh9WMrEAgEYBiG49L6GqseX4ubbrkNDQ274ff74ff7LVvBesQN3S29bS80TYOmafB4PBg5YjgWnXEavnvOWS",
  "grK00pjXWrAUgA3gAwxWrdFWkPR11SMB/AU0IIU5IkfuDAQRw/aZpTwpZlGa2tbRg/fhzuu+cujD32mJSNCiHAZdlZfOfO93",
  "DLz+/EE+uehNfrRVZWVq/G1OZi0kiaThNG8rQ79dA0FoshGouhvKwU/3bhBbjsxz9EMJjTnQh2+8wCxtjfux+Qsj4aI14CUG",
  "8YhinLMl9yxdX4zf2/RZFVy1cUBR0dnVDVLJx37tk45+wzMWb0KMhWENPW3o4333wLj61ajXVPPoWOzk7khkI9wmPbTUqMQU",
  "voiMWi0A0DMufw+7Pg8SiIRqOIxzUAgM/ntSSHO+5USyTQ0d6OESOG447bbsa8k+fa/7PBb2SMzUjXKHHU43Er4pJaW1tZ3Y",
  "w5+PTTA1DV5LG4LMswDAPtHR0IqCqGDRuKqnAYmhbHrobd2P/JpzAMAznZ2ZBluUfubnO7qyvZPFVaUoLxx43FxAknYNTIEQ",
  "iHKxFQVRw63IQ9ez7E29u2Y9PmLdjV0IBIJIpAQIXP53NsT5JQcfzsxmW49uqlZKlwn8fj/WqQMAzTkGUub37jH5h/yukQQs",
  "Dn8yWtttUXYJomYvG4U+Hxer3wer1ODtA9qeGcJ7vGNA3Hjz8O5593LubPm4vysrK+T5cMA9t3vIPH1zyB1WvW4eOPP0YgEI",
  "DP57UaMzTk5eZi29ZNRiAQkAVwG8+kQaK3FhlbFTa8uBFnnn0eIpGoU+Vxt664y1vdDZttKOPxODo7u3DsMaNx5RWX4dvfOt",
  "1RHXeLjb0WETmBvNvANTUdwcMrHsV9DzyI/fs/gapmIaHrKCosNPd9sJMzxrZs3bq1bvz48Zm1yHRvkhJCvClJUp5hGEKWZW",
  "nb9h348ZKl2PzGFvj8fmRZEVtvvpiIkhISiyFunect/v5F+NHiS5Bl9QL205+ndJLZxDh4sBG3/vxOvPLqazAFiaLCAunpv6",
  "1t8fl8R22S6m/DxEy761vXDdM+hvrtg3+guvrZlFtYToo/SNybk/ZS/CHKL66kaTNPol/dfS8dOnS4RwQ5oH4BIchqYSAior",
  "immfF43NQ0Tbdb64/WPzygRknTNInbwT6Ad97ZiXd37sRnBw4iEomCQM7CgUAA4XAlxoweheHDhmYawfW7JmEKIWTOM26UzL",
  "RPeBERJbq3z2Qy9Mxi+Ew6yMk0zYSu64sy6RfOpPJjS8IsACsBFAIwiEi2u0Z7KyPYBm2AWdzRhmEd8jYB+A5jbMMXxvkvsl",
  "3+S3pv4Ktrl//6hYmvX5n5+qWp/5evzf3Lvzj59auzX788/S/++vx/A2UA0pvbEPR7AAAAAElFTkSuQmCC",
].join("");


const CWS_HTML_FORMAT = "CWS_HTML";


const CWS_SCHEMA_VERSION = 1;


const CWS_AI_GUIDE_URL = "https://chtec.co.jp/cws/ai/cws-html-guide-v1.json";


const CWS_MODEL_SCHEMA_URL = "https://chtec.co.jp/cws/schema/cws-html-workbook-model-v1.schema.json";


const CWS_WEB_APP_URL = "https://chtec.co.jp/cws/";


const SERVER_VERSION_INFO_PATH = "download/version.json";


const SERVER_DOWNLOAD_PAGE_PATH = "https://chtec.co.jp/cws/manual/";


const MERGE_CONTENT_WARNING_MESSAGE = "セルを結合すると、左上の値のみが保持され、他のセルの値は破棄されます。";


const CWS_CLIPBOARD_RANGE_MIME = "application/x-cht-websheet-range";


const CWS_CLIPBOARD_RANGE_MARKER = "Cht WebSheet Range";


const CWS_CLIPBOARD_DATA_MIME = "application/x-cht-websheet-clipboard+json";


const CWS_CLIPBOARD_DATA_WEB_MIME = `web ${CWS_CLIPBOARD_DATA_MIME}`;


const CWS_CLIPBOARD_HTML_MARKER = "data-cht-websheet-clipboard";


const DIALOG_CHROME_ROOT_SELECTOR = [
  ".excel-mini-dialog[role='dialog']",
  ".format-dialog[role='dialog']",
  ".cell-structure-dialog[role='dialog']",
  ".message-box[role='alertdialog']",
  ".zoom-dialog[role='dialog']",
].join(",");


const DIALOG_CHROME_TITLE_SELECTOR = [
  ".excel-mini-titlebar",
  ".format-dialog-titlebar",
  ".cell-structure-titlebar",
  ".message-box-title",
  ".zoom-dialog-title",
].join(",");


const STANDALONE_RUNTIME_BASE_URL_STORAGE_KEY = "websheet.standaloneRuntimeBaseUrl";


const STANDALONE_REMOTE_RUNTIME_BASE_URL_STORAGE_KEY = "websheet.remoteRuntimeBaseUrl";


const AUTO_STANDALONE_RUNTIME_BASE_URL = "__auto__";


const RECENT_DATA_SOURCES_STORAGE_KEY = "websheet.recentDataSources";


const RECENT_DATA_SOURCES_LIMIT = 12;


const RECENT_DATA_SOURCE_TEXT_LIMIT = 220000;


const DEFAULT_STANDALONE_RUNTIME_BASE_URL =
  typeof __WEBSHEET_DEFAULT_RUNTIME_BASE_URL__ === "string"
    ? __WEBSHEET_DEFAULT_RUNTIME_BASE_URL__
    : `./websheet-runtime/${STANDALONE_RUNTIME_CHANNEL}`;


const DEFAULT_STANDALONE_REMOTE_RUNTIME_BASE_URL = `https://chtec.co.jp/cws/download/runtime/${STANDALONE_RUNTIME_CHANNEL}`;


const APP_TITLE = "WebSheet HTML Converter";


const HISTORY_LIMIT = 100;


const DATA_OPERATION_CELL_LIMIT = 100000;


const AUTO_FILTER_VALUE_LIMIT = 1000;


const AUTO_FILTER_HEADER_SCAN_LIMIT = 500;


const STYLE_INTERN_POOL_LIMIT = 20000;


const PERSISTENCE_DB_NAME = "cht-websheet";


const PERSISTENCE_DB_VERSION = 1;


const PERSISTENCE_STORE_NAME = "snapshots";


const PERSISTENCE_CURRENT_KEY = "current-workbook";


const PERSISTENCE_SNAPSHOT_VERSION = 1;


const PENDING_FILE_LOAD_RECOVERY_STORAGE_KEY = "websheet.pendingFileLoadRecovery.v1";


const WORKBOOK_OPEN_TRANSFER_DB_NAME = "cht-websheet-open-transfers";


const WORKBOOK_OPEN_TRANSFER_DB_VERSION = 1;


const WORKBOOK_OPEN_TRANSFER_STORE_NAME = "files";


const WORKBOOK_OPEN_TRANSFER_QUERY_PARAM = "websheetOpenToken";


const WORKBOOK_OPEN_TRANSFER_MAX_AGE_MS = 10 * 60 * 1000;


const WORKBOOK_MODEL_VERSION = 8;


const PERSISTENCE_SAVE_DEBOUNCE_MS = 1500;


const PERSISTENCE_IDLE_TIMEOUT_MS = 3000;


const EXCELJS_IMPORT_IGNORE_NODES = ["drawing", "picture", "tableParts"];


const DEFAULT_MACRO_SECURITY_LEVEL = "disableWithNotification";


const WEBSHEET_FORMULA_NA_SENTINEL = "__WEBSHEET_FORMULA_NA__";



let ExcelJS = ExcelJSModule?.default || ExcelJSModule;


let JSZip = JSZipModule?.default || JSZipModule;


let XLSX = XLSXModule;


let excelJsModulePromise = null;


let jsZipModulePromise = null;


let xlsxModulePromise = null;


const worksheetImageRenderObjectUrls = new Map();


const WORKSHEET_IMAGE_OBJECT_URL_MIN_DATA_URL_LENGTH = 4096;


const WORKSHEET_OBJECT_LAYER_BASE_Z = 30;


const WORKSHEET_OBJECT_LAYER_STEP_Z = 3;


const WORKSHEET_OBJECT_SELECTION_Z_OFFSET = 1;


const SHAPE_TEXT_EDITOR_Z_INDEX = 10000;



async function loadExcelJsModule() {
  if (ExcelJS) return ExcelJS;
  excelJsModulePromise = excelJsModulePromise || Promise.resolve(ExcelJSModule).then((module) => {
    ExcelJS = module?.default || module;
    return ExcelJS;
  });
  return excelJsModulePromise;
}



async function loadJsZipModule() {
  if (JSZip) return JSZip;
  jsZipModulePromise = jsZipModulePromise || Promise.resolve(JSZipModule).then((module) => {
    JSZip = module?.default || module;
    return JSZip;
  });
  return jsZipModulePromise;
}



async function loadXlsxModule() {
  if (XLSX) return XLSX;
  xlsxModulePromise = xlsxModulePromise || Promise.resolve(XLSXModule).then((module) => {
    XLSX = module;
    return XLSX;
  });
  return xlsxModulePromise;
}



warmWorkbookIoModules();



const WEBSHEET_ROMAN_VALUES = [
  [1000, "M"],
  [900, "CM"],
  [500, "D"],
  [400, "CD"],
  [100, "C"],
  [90, "XC"],
  [50, "L"],
  [40, "XL"],
  [10, "X"],
  [9, "IX"],
  [5, "V"],
  [4, "IV"],
  [1, "I"],
];



const WEBSHEET_CONVERT_UNITS = new Map([
  ["m", ["length", 1]],
  ["meter", ["length", 1]],
  ["metre", ["length", 1]],
  ["km", ["length", 1000]],
  ["cm", ["length", 0.01]],
  ["mm", ["length", 0.001]],
  ["um", ["length", 0.000001]],
  ["µm", ["length", 0.000001]],
  ["nm", ["length", 0.000000001]],
  ["in", ["length", 0.0254]],
  ["ft", ["length", 0.3048]],
  ["yd", ["length", 0.9144]],
  ["mi", ["length", 1609.344]],
  ["nmi", ["length", 1852]],
  ["ang", ["length", 1e-10]],
  ["pica", ["length", 0.004233333333333333]],
  ["g", ["mass", 1]],
  ["gram", ["mass", 1]],
  ["kg", ["mass", 1000]],
  ["mg", ["mass", 0.001]],
  ["lbm", ["mass", 453.59237]],
  ["ozm", ["mass", 28.349523125]],
  ["slug", ["mass", 14593.90294]],
  ["ton", ["mass", 907184.74]],
  ["u", ["mass", 1.6605390666e-24]],
  ["s", ["time", 1]],
  ["sec", ["time", 1]],
  ["mn", ["time", 60]],
  ["min", ["time", 60]],
  ["hr", ["time", 3600]],
  ["day", ["time", 86400]],
  ["yr", ["time", 31557600]],
  ["pa", ["pressure", 1]],
  ["atm", ["pressure", 101325]],
  ["mmhg", ["pressure", 133.32236842105263]],
  ["n", ["force", 1]],
  ["dyn", ["force", 0.00001]],
  ["lbf", ["force", 4.4482216152605]],
  ["j", ["energy", 1]],
  ["ev", ["energy", 1.602176634e-19]],
  ["cal", ["energy", 4.184]],
  ["btu", ["energy", 1055.05585262]],
  ["w", ["power", 1]],
  ["hp", ["power", 745.6998715822702]],
  ["l", ["volume", 1]],
  ["lt", ["volume", 1]],
  ["liter", ["volume", 1]],
  ["litre", ["volume", 1]],
  ["ml", ["volume", 0.001]],
  ["m3", ["volume", 1000]],
  ["gal", ["volume", 3.785411784]],
  ["qt", ["volume", 0.946352946]],
  ["pt", ["volume", 0.473176473]],
  ["cup", ["volume", 0.2365882365]],
  ["floz", ["volume", 0.0295735295625]],
  ["m2", ["area", 1]],
  ["cm2", ["area", 0.0001]],
  ["mm2", ["area", 0.000001]],
  ["ft2", ["area", 0.09290304]],
  ["in2", ["area", 0.00064516]],
  ["yd2", ["area", 0.83612736]],
  ["acre", ["area", 4046.8564224]],
]);



const WebSheetFormulaPlugin = createWebSheetFormulaPluginClass();



const WEBSHEET_FORMULA_RANGE = { argumentType: FunctionArgumentType.RANGE };


const WEBSHEET_FORMULA_ANY = { argumentType: FunctionArgumentType.ANY };


const WEBSHEET_FORMULA_NUMBER = { argumentType: FunctionArgumentType.NUMBER };


const WEBSHEET_FORMULA_INTEGER = { argumentType: FunctionArgumentType.INTEGER };


const WEBSHEET_FORMULA_STRING = { argumentType: FunctionArgumentType.STRING };


const WEBSHEET_FORMULA_BOOLEAN = { argumentType: FunctionArgumentType.BOOLEAN };


const WEBSHEET_FORMULA_SCALAR = { argumentType: FunctionArgumentType.SCALAR };



WebSheetFormulaPlugin.implementedFunctions = {
  TEXTJOIN: {
    method: "textjoin",
    parameters: [WEBSHEET_FORMULA_STRING, WEBSHEET_FORMULA_BOOLEAN, WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  CONCAT: {
    method: "concat",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  CONCATENATE: {
    method: "concatenate",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  REPLACE: {
    method: "replaceText",
    parameters: [WEBSHEET_FORMULA_STRING, WEBSHEET_FORMULA_INTEGER, WEBSHEET_FORMULA_INTEGER, WEBSHEET_FORMULA_STRING],
  },
  PROPER: {
    method: "proper",
    parameters: [WEBSHEET_FORMULA_STRING],
  },
  EXACT: {
    method: "exact",
    parameters: [WEBSHEET_FORMULA_STRING, WEBSHEET_FORMULA_STRING],
  },
  CHAR: {
    method: "char",
    parameters: [WEBSHEET_FORMULA_INTEGER],
  },
  CODE: {
    method: "code",
    parameters: [WEBSHEET_FORMULA_STRING],
  },
  CLEAN: {
    method: "clean",
    parameters: [WEBSHEET_FORMULA_STRING],
  },
  REPT: {
    method: "rept",
    parameters: [WEBSHEET_FORMULA_STRING, WEBSHEET_FORMULA_INTEGER],
  },
  TEXTSPLIT: {
    method: "textsplit",
    sizeOfResultArrayMethod: "textsplitArraySize",
    parameters: [
      WEBSHEET_FORMULA_ANY,
      WEBSHEET_FORMULA_ANY,
      { ...WEBSHEET_FORMULA_ANY, optionalArg: true },
      { ...WEBSHEET_FORMULA_BOOLEAN, defaultValue: false },
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
      { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL },
    ],
    vectorizationForbidden: true,
  },
  TEXTBEFORE: {
    method: "textbefore",
    parameters: [
      WEBSHEET_FORMULA_ANY,
      WEBSHEET_FORMULA_ANY,
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 1 },
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
      { ...WEBSHEET_FORMULA_BOOLEAN, defaultValue: false },
      { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL },
    ],
  },
  TEXTAFTER: {
    method: "textafter",
    parameters: [
      WEBSHEET_FORMULA_ANY,
      WEBSHEET_FORMULA_ANY,
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 1 },
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
      { ...WEBSHEET_FORMULA_BOOLEAN, defaultValue: false },
      { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL },
    ],
  },
  AGGREGATE: {
    method: "aggregate",
    parameters: [WEBSHEET_FORMULA_INTEGER, WEBSHEET_FORMULA_INTEGER, WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  IFS: {
    method: "ifs",
    parameters: [WEBSHEET_FORMULA_BOOLEAN, WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 2,
  },
  SWITCH: {
    method: "switchFunction",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  XOR: {
    method: "xor",
    parameters: [WEBSHEET_FORMULA_BOOLEAN],
    repeatLastArgs: 1,
  },
  LET: {
    method: "letFunction",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 2,
  },
  LAMBDA: {
    method: "lambda",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  SUMPRODUCT: {
    method: "sumproduct",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  COUNTBLANK: {
    method: "countblank",
    parameters: [WEBSHEET_FORMULA_RANGE],
  },
  COUNTIF: {
    method: "countif",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_SCALAR],
  },
  COUNTIFS: {
    method: "countifs",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_SCALAR],
    repeatLastArgs: 2,
  },
  SUMIF: {
    method: "sumif",
    parameters: [
      WEBSHEET_FORMULA_RANGE,
      WEBSHEET_FORMULA_SCALAR,
      { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL },
    ],
  },
  SUMIFS: {
    method: "sumifs",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_SCALAR],
    repeatLastArgs: 2,
  },
  AVERAGEIF: {
    method: "averageif",
    parameters: [
      WEBSHEET_FORMULA_RANGE,
      WEBSHEET_FORMULA_SCALAR,
      { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL },
    ],
  },
  MAKEARRAY: {
    method: "makearray",
    sizeOfResultArrayMethod: "makearrayArraySize",
    parameters: [WEBSHEET_FORMULA_INTEGER, WEBSHEET_FORMULA_INTEGER, WEBSHEET_FORMULA_ANY],
    vectorizationForbidden: true,
  },
  BYROW: {
    method: "byrow",
    sizeOfResultArrayMethod: "byrowArraySize",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_ANY],
    vectorizationForbidden: true,
  },
  BYCOL: {
    method: "bycol",
    sizeOfResultArrayMethod: "bycolArraySize",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_ANY],
    vectorizationForbidden: true,
  },
  MAP: {
    method: "mapFunction",
    sizeOfResultArrayMethod: "dynamicArraySize",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
    vectorizationForbidden: true,
  },
  REDUCE: {
    method: "reduce",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_ANY],
  },
  SCAN: {
    method: "scan",
    sizeOfResultArrayMethod: "scanArraySize",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_ANY],
    vectorizationForbidden: true,
  },
  AVERAGEIFS: {
    method: "averageifs",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_SCALAR],
    repeatLastArgs: 2,
  },
  AVERAGEA: {
    method: "averagea",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  MAXA: {
    method: "maxa",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  MINA: {
    method: "mina",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  "STDEV.S": {
    method: "stdevS",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  "STDEV.P": {
    method: "stdevP",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  STDEV: {
    method: "stdev",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  STDEVP: {
    method: "stdevp",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  STDEVA: {
    method: "stdeva",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  STDEVPA: {
    method: "stdevpa",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  "VAR.S": {
    method: "varS",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  "VAR.P": {
    method: "varP",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  VAR: {
    method: "varLegacy",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  VARP: {
    method: "varp",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  VARA: {
    method: "vara",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  VARPA: {
    method: "varpa",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  CORREL: {
    method: "correl",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY],
  },
  "COVARIANCE.S": {
    method: "covarianceS",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY],
  },
  "COVARIANCE.P": {
    method: "covarianceP",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY],
  },
  COVAR: {
    method: "covar",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY],
  },
  SLOPE: {
    method: "slope",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY],
  },
  LINEST: {
    method: "linest",
    sizeOfResultArrayMethod: "linestArraySize",
    parameters: [
      WEBSHEET_FORMULA_RANGE,
      { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL },
      { ...WEBSHEET_FORMULA_BOOLEAN, defaultValue: true },
      { ...WEBSHEET_FORMULA_BOOLEAN, defaultValue: false },
    ],
    vectorizationForbidden: true,
  },
  LOGEST: {
    method: "logest",
    sizeOfResultArrayMethod: "linestArraySize",
    parameters: [
      WEBSHEET_FORMULA_RANGE,
      { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL },
      { ...WEBSHEET_FORMULA_BOOLEAN, defaultValue: true },
      { ...WEBSHEET_FORMULA_BOOLEAN, defaultValue: false },
    ],
    vectorizationForbidden: true,
  },
  AVEDEV: {
    method: "avedev",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  DEVSQ: {
    method: "devsq",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  GEOMEAN: {
    method: "geomean",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  HARMEAN: {
    method: "harmean",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  PEARSON: {
    method: "pearson",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY],
  },
  RSQ: {
    method: "rsq",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY],
  },
  STEYX: {
    method: "steyx",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY],
  },
  SKEW: {
    method: "skew",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  "SKEW.P": {
    method: "skewP",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  "NORM.DIST": {
    method: "normDist",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_BOOLEAN],
  },
  NORMDIST: {
    method: "normDistLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_BOOLEAN],
  },
  "NORM.S.DIST": {
    method: "normSDist",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_BOOLEAN],
  },
  NORMSDIST: {
    method: "normSDistLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  "NORM.INV": {
    method: "normInv",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  NORMINV: {
    method: "normInvLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  "NORM.S.INV": {
    method: "normSInv",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  NORMSINV: {
    method: "normSInvLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  STANDARDIZE: {
    method: "standardize",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  TRIMMEAN: {
    method: "trimmean",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_NUMBER],
  },
  "EXPON.DIST": {
    method: "exponDist",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_BOOLEAN],
  },
  EXPONDIST: {
    method: "exponDistLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_BOOLEAN],
  },
  "POISSON.DIST": {
    method: "poissonDist",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_BOOLEAN],
  },
  POISSON: {
    method: "poissonDistLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_BOOLEAN],
  },
  "BINOM.DIST": {
    method: "binomDist",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_BOOLEAN],
  },
  "BINOM.DIST.RANGE": {
    method: "binomDistRange",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL },
    ],
  },
  BINOMDIST: {
    method: "binomDistLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_BOOLEAN],
  },
  "BINOM.INV": {
    method: "binomInv",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  CRITBINOM: {
    method: "critbinom",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  FISHER: {
    method: "fisher",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  FISHERINV: {
    method: "fisherinv",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  GAMMA: {
    method: "gamma",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  GAMMALN: {
    method: "gammaln",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  "GAMMALN.PRECISE": {
    method: "gammalnPrecise",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  "GAMMA.DIST": {
    method: "gammaDist",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_BOOLEAN],
  },
  GAMMADIST: {
    method: "gammaDistLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_BOOLEAN],
  },
  "GAMMA.INV": {
    method: "gammaInv",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  GAMMAINV: {
    method: "gammaInvLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  "BETA.DIST": {
    method: "betaDist",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_BOOLEAN,
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 0 },
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 1 },
    ],
  },
  BETADIST: {
    method: "betaDistLegacy",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 0 },
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 1 },
    ],
  },
  "BETA.INV": {
    method: "betaInv",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 0 },
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 1 },
    ],
  },
  BETAINV: {
    method: "betaInvLegacy",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 0 },
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 1 },
    ],
  },
  "CHISQ.DIST": {
    method: "chisqDist",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_BOOLEAN],
  },
  "CHISQ.INV": {
    method: "chisqInv",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  CHIDIST: {
    method: "chiDistLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  CHIDISTRT: {
    method: "chiDistLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  CHIINV: {
    method: "chiInvLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  CHIINVRT: {
    method: "chiInvLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  "CHISQ.TEST": {
    method: "chisqTest",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY],
  },
  CHITEST: {
    method: "chitestLegacy",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY],
  },
  "WEIBULL.DIST": {
    method: "weibullDist",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_BOOLEAN],
  },
  WEIBULL: {
    method: "weibullDistLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_BOOLEAN],
  },
  WEIBULLDIST: {
    method: "weibullDistLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_BOOLEAN],
  },
  "HYPGEOM.DIST": {
    method: "hypgeomDist",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_BOOLEAN],
  },
  HYPGEOMDIST: {
    method: "hypgeomDistLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  "CONFIDENCE.NORM": {
    method: "confidenceNorm",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  "CONFIDENCE.T": {
    method: "confidenceT",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  CONFIDENCE: {
    method: "confidenceLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  "F.DIST": {
    method: "fDist",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_BOOLEAN],
  },
  "F.DIST.RT": {
    method: "fDistRt",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  "F.INV": {
    method: "fInv",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  "F.INV.RT": {
    method: "fInvRt",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  FDIST: {
    method: "fDistLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  FDISTRT: {
    method: "fDistLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  FINV: {
    method: "fInvLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  FINVRT: {
    method: "fInvLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  "F.TEST": {
    method: "fTest",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY],
  },
  FTEST: {
    method: "fTestLegacy",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY],
  },
  "Z.TEST": {
    method: "zTest",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_NUMBER, { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL }],
  },
  ZTEST: {
    method: "zTestLegacy",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_NUMBER, { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL }],
  },
  "T.TEST": {
    method: "tTest",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_INTEGER, WEBSHEET_FORMULA_INTEGER],
  },
  TTEST: {
    method: "ttestLegacy",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_INTEGER, WEBSHEET_FORMULA_INTEGER],
  },
  "T.DIST": {
    method: "tDist",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_BOOLEAN],
  },
  "T.DIST.RT": {
    method: "tDistRt",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  "T.DIST.2T": {
    method: "tDist2t",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  "T.INV": {
    method: "tInv",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  "T.INV.2T": {
    method: "tInv2t",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  TDIST: {
    method: "tDistLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  TDISTRT: {
    method: "tDistRtLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  TDIST2T: {
    method: "tDist2tLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  TINV: {
    method: "tInvLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  TINV2T: {
    method: "tInvLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  "LOGNORM.DIST": {
    method: "lognormDist",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_BOOLEAN],
  },
  LOGNORMDIST: {
    method: "lognormDistLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  "LOGNORM.INV": {
    method: "lognormInv",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  LOGINV: {
    method: "lognormInvLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  LOGNORMINV: {
    method: "lognormInvLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  "NEGBINOM.DIST": {
    method: "negbinomDist",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_BOOLEAN],
  },
  NEGBINOMDIST: {
    method: "negbinomDistLegacy",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  ERF: {
    method: "erf",
    parameters: [WEBSHEET_FORMULA_NUMBER, { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL }],
  },
  ERFC: {
    method: "erfc",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  "ERF.PRECISE": {
    method: "erfPrecise",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  "ERFC.PRECISE": {
    method: "erfcPrecise",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  PHI: {
    method: "phi",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  GAUSS: {
    method: "gauss",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  MOD: {
    method: "mod",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  POWER: {
    method: "power",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  SQRT: {
    method: "sqrt",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  SQRTPI: {
    method: "sqrtpi",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  LN: {
    method: "ln",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  LOG: {
    method: "log",
    parameters: [WEBSHEET_FORMULA_NUMBER, { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL }],
  },
  LOG10: {
    method: "log10",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  EXP: {
    method: "exp",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  FACT: {
    method: "fact",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  FACTDOUBLE: {
    method: "factdouble",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  COMBIN: {
    method: "combin",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  COMBINA: {
    method: "combina",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  PERMUT: {
    method: "permut",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  PERMUTATIONA: {
    method: "permutationa",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  MULTINOMIAL: {
    method: "multinomial",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  GCD: {
    method: "gcd",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  LCM: {
    method: "lcm",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  PI: {
    method: "pi",
    parameters: [],
  },
  SIN: {
    method: "sin",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  COS: {
    method: "cos",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  TAN: {
    method: "tan",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  ASIN: {
    method: "asin",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  ACOS: {
    method: "acos",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  ATAN: {
    method: "atan",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  ATAN2: {
    method: "atan2",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  DEGREES: {
    method: "degrees",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  RADIANS: {
    method: "radians",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  SINH: {
    method: "sinh",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  COSH: {
    method: "cosh",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  TANH: {
    method: "tanh",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  ASINH: {
    method: "asinh",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  ACOSH: {
    method: "acosh",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  ATANH: {
    method: "atanh",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  COT: {
    method: "cot",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  SEC: {
    method: "sec",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  CSC: {
    method: "csc",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  ACOT: {
    method: "acot",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  COTH: {
    method: "coth",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  SECH: {
    method: "sech",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  CSCH: {
    method: "csch",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  SERIESSUM: {
    method: "seriessum",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_ANY],
  },
  SUMSQ: {
    method: "sumsq",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  SUMX2MY2: {
    method: "sumx2my2",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY],
  },
  SUMX2PY2: {
    method: "sumx2py2",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY],
  },
  SUMXMY2: {
    method: "sumxmy2",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY],
  },
  EVEN: {
    method: "even",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  ODD: {
    method: "odd",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  "ISO.CEILING": {
    method: "isoCeiling",
    parameters: [WEBSHEET_FORMULA_NUMBER, { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 1 }],
  },
  MAXIFS: {
    method: "maxifs",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_SCALAR],
    repeatLastArgs: 2,
  },
  MINIFS: {
    method: "minifs",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_SCALAR],
    repeatLastArgs: 2,
  },
  "MODE.SNGL": {
    method: "modeSngl",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  MODE: {
    method: "modeSngl",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  "MODE.MULT": {
    method: "modeMult",
    sizeOfResultArrayMethod: "modeMultArraySize",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
    vectorizationForbidden: true,
  },
  KURT: {
    method: "kurt",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  "RANK.EQ": {
    method: "rankEq",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_RANGE, { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 0 }],
  },
  RANK: {
    method: "rankEq",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_RANGE, { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 0 }],
  },
  "RANK.AVG": {
    method: "rankAvg",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_RANGE, { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 0 }],
  },
  "PERCENTILE.INC": {
    method: "percentileInc",
    parameters: [WEBSHEET_FORMULA_RANGE, { ...WEBSHEET_FORMULA_NUMBER, minValue: 0, maxValue: 1 }],
  },
  PERCENTILE: {
    method: "percentileInc",
    parameters: [WEBSHEET_FORMULA_RANGE, { ...WEBSHEET_FORMULA_NUMBER, minValue: 0, maxValue: 1 }],
  },
  "PERCENTILE.EXC": {
    method: "percentileExc",
    parameters: [WEBSHEET_FORMULA_RANGE, { ...WEBSHEET_FORMULA_NUMBER, minValue: 0, maxValue: 1 }],
  },
  "PERCENTRANK.INC": {
    method: "percentRankInc",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_NUMBER, { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 3 }],
  },
  "PERCENTRANK.EXC": {
    method: "percentRankExc",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_NUMBER, { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 3 }],
  },
  "QUARTILE.INC": {
    method: "quartileInc",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_INTEGER],
  },
  QUARTILE: {
    method: "quartileInc",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_INTEGER],
  },
  "QUARTILE.EXC": {
    method: "quartileExc",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_INTEGER],
  },
  PROB: {
    method: "prob",
    parameters: [
      WEBSHEET_FORMULA_RANGE,
      WEBSHEET_FORMULA_RANGE,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL },
    ],
  },
  FREQUENCY: {
    method: "frequency",
    sizeOfResultArrayMethod: "frequencyArraySize",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_RANGE],
    vectorizationForbidden: true,
  },
  INTERCEPT: {
    method: "intercept",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_RANGE],
  },
  "FORECAST.LINEAR": {
    method: "forecastLinear",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_RANGE],
  },
  FORECAST: {
    method: "forecastLinear",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_RANGE],
  },
  TREND: {
    method: "trend",
    sizeOfResultArrayMethod: "trendArraySize",
    parameters: [
      WEBSHEET_FORMULA_RANGE,
      { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL },
      { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL },
      { ...WEBSHEET_FORMULA_BOOLEAN, defaultValue: true },
    ],
    vectorizationForbidden: true,
  },
  GROWTH: {
    method: "growth",
    sizeOfResultArrayMethod: "trendArraySize",
    parameters: [
      WEBSHEET_FORMULA_RANGE,
      { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL },
      { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL },
      { ...WEBSHEET_FORMULA_BOOLEAN, defaultValue: true },
    ],
    vectorizationForbidden: true,
  },
  DOLLAR: {
    method: "dollar",
    parameters: [WEBSHEET_FORMULA_NUMBER, { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 2 }],
  },
  FIXED: {
    method: "fixed",
    parameters: [WEBSHEET_FORMULA_NUMBER, { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 2 }, { ...WEBSHEET_FORMULA_BOOLEAN, defaultValue: false }],
  },
  NUMBERVALUE: {
    method: "numbervalue",
    parameters: [WEBSHEET_FORMULA_STRING, { ...WEBSHEET_FORMULA_STRING, defaultValue: "." }, { ...WEBSHEET_FORMULA_STRING, defaultValue: "," }],
  },
  "ERROR.TYPE": {
    method: "errorType",
    parameters: [WEBSHEET_FORMULA_SCALAR],
  },
  ISBLANK: {
    method: "isblank",
    parameters: [WEBSHEET_FORMULA_SCALAR],
  },
  ISERROR: {
    method: "iserror",
    parameters: [WEBSHEET_FORMULA_SCALAR],
  },
  ISERR: {
    method: "iserr",
    parameters: [WEBSHEET_FORMULA_SCALAR],
  },
  ISNA: {
    method: "isna",
    parameters: [WEBSHEET_FORMULA_SCALAR],
  },
  ISNUMBER: {
    method: "isnumber",
    parameters: [WEBSHEET_FORMULA_SCALAR],
  },
  ISTEXT: {
    method: "istext",
    parameters: [WEBSHEET_FORMULA_SCALAR],
  },
  ISNONTEXT: {
    method: "isnontext",
    parameters: [WEBSHEET_FORMULA_SCALAR],
  },
  ISLOGICAL: {
    method: "islogical",
    parameters: [WEBSHEET_FORMULA_SCALAR],
  },
  ISEVEN: {
    method: "iseven",
    parameters: [WEBSHEET_FORMULA_SCALAR],
  },
  ISODD: {
    method: "isodd",
    parameters: [WEBSHEET_FORMULA_SCALAR],
  },
  N: {
    method: "nFunction",
    parameters: [WEBSHEET_FORMULA_SCALAR],
  },
  NA: {
    method: "na",
    parameters: [],
  },
  TYPE: {
    method: "type",
    parameters: [WEBSHEET_FORMULA_SCALAR],
  },
  ISREF: {
    method: "isref",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  ISFORMULA: {
    method: "isformula",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  FORMULATEXT: {
    method: "formulatext",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  SHEET: {
    method: "sheet",
    parameters: [{ ...WEBSHEET_FORMULA_ANY, optionalArg: true }],
  },
  SHEETS: {
    method: "sheets",
    parameters: [{ ...WEBSHEET_FORMULA_ANY, optionalArg: true }],
  },
  CELL: {
    method: "cell",
    parameters: [WEBSHEET_FORMULA_STRING, { ...WEBSHEET_FORMULA_ANY, optionalArg: true }],
  },
  INFO: {
    method: "info",
    parameters: [WEBSHEET_FORMULA_STRING],
  },
  ROW: {
    method: "row",
    parameters: [{ ...WEBSHEET_FORMULA_ANY, optionalArg: true }],
  },
  COLUMN: {
    method: "column",
    parameters: [{ ...WEBSHEET_FORMULA_ANY, optionalArg: true }],
  },
  ROWS: {
    method: "rows",
    parameters: [WEBSHEET_FORMULA_RANGE],
  },
  COLUMNS: {
    method: "columns",
    parameters: [WEBSHEET_FORMULA_RANGE],
  },
  HYPERLINK: {
    method: "hyperlink",
    parameters: [WEBSHEET_FORMULA_STRING, { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL }],
  },
  T: {
    method: "tFunction",
    parameters: [WEBSHEET_FORMULA_SCALAR],
  },
  LOOKUP: {
    method: "lookup",
    parameters: [WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_RANGE, { ...WEBSHEET_FORMULA_ANY, optionalArg: true }],
  },
  MMULT: {
    method: "mmult",
    sizeOfResultArrayMethod: "mmultArraySize",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_RANGE],
    vectorizationForbidden: true,
  },
  MDETERM: {
    method: "mdeterm",
    parameters: [WEBSHEET_FORMULA_RANGE],
  },
  MINVERSE: {
    method: "minverse",
    sizeOfResultArrayMethod: "minverseArraySize",
    parameters: [WEBSHEET_FORMULA_RANGE],
    vectorizationForbidden: true,
  },
  MUNIT: {
    method: "munit",
    sizeOfResultArrayMethod: "munitArraySize",
    parameters: [WEBSHEET_FORMULA_INTEGER],
    vectorizationForbidden: true,
  },
  XLOOKUP: {
    method: "xlookup",
    sizeOfResultArrayMethod: "xlookupArraySize",
    parameters: [
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_RANGE,
      WEBSHEET_FORMULA_RANGE,
      { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL },
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 1 },
    ],
    vectorizationForbidden: true,
  },
  XMATCH: {
    method: "xmatch",
    parameters: [
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_RANGE,
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 1 },
    ],
  },
  FILTER: {
    method: "filter",
    sizeOfResultArrayMethod: "dynamicArraySize",
    parameters: [
      WEBSHEET_FORMULA_RANGE,
      WEBSHEET_FORMULA_RANGE,
      { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL },
    ],
    vectorizationForbidden: true,
  },
  SORT: {
    method: "sort",
    sizeOfResultArrayMethod: "sortArraySize",
    parameters: [
      WEBSHEET_FORMULA_RANGE,
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 1 },
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 1 },
      { ...WEBSHEET_FORMULA_BOOLEAN, defaultValue: false },
    ],
    vectorizationForbidden: true,
  },
  SORTBY: {
    method: "sortby",
    sizeOfResultArrayMethod: "sortArraySize",
    parameters: [
      WEBSHEET_FORMULA_RANGE,
      WEBSHEET_FORMULA_RANGE,
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 1 },
    ],
    repeatLastArgs: 2,
    vectorizationForbidden: true,
  },
  UNIQUE: {
    method: "unique",
    sizeOfResultArrayMethod: "uniqueArraySize",
    parameters: [
      WEBSHEET_FORMULA_RANGE,
      { ...WEBSHEET_FORMULA_BOOLEAN, defaultValue: false },
      { ...WEBSHEET_FORMULA_BOOLEAN, defaultValue: false },
    ],
    vectorizationForbidden: true,
  },
  SEQUENCE: {
    method: "sequence",
    sizeOfResultArrayMethod: "sequenceArraySize",
    parameters: [
      WEBSHEET_FORMULA_INTEGER,
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 1 },
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 1 },
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 1 },
    ],
    vectorizationForbidden: true,
  },
  INDIRECT: {
    method: "indirect",
    sizeOfResultArrayMethod: "indirectArraySize",
    parameters: [
      WEBSHEET_FORMULA_STRING,
      { ...WEBSHEET_FORMULA_BOOLEAN, defaultValue: true },
    ],
    vectorizationForbidden: true,
  },
  TAKE: {
    method: "take",
    sizeOfResultArrayMethod: "takeArraySize",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_INTEGER, { ...WEBSHEET_FORMULA_INTEGER, optionalArg: true }],
    vectorizationForbidden: true,
  },
  DROP: {
    method: "drop",
    sizeOfResultArrayMethod: "dropArraySize",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_INTEGER, { ...WEBSHEET_FORMULA_INTEGER, optionalArg: true }],
    vectorizationForbidden: true,
  },
  VSTACK: {
    method: "vstack",
    sizeOfResultArrayMethod: "vstackArraySize",
    parameters: [WEBSHEET_FORMULA_RANGE],
    repeatLastArgs: 1,
    vectorizationForbidden: true,
  },
  HSTACK: {
    method: "hstack",
    sizeOfResultArrayMethod: "hstackArraySize",
    parameters: [WEBSHEET_FORMULA_RANGE],
    repeatLastArgs: 1,
    vectorizationForbidden: true,
  },
  CHOOSEROWS: {
    method: "chooserows",
    sizeOfResultArrayMethod: "chooserowsArraySize",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_INTEGER],
    repeatLastArgs: 1,
    vectorizationForbidden: true,
  },
  CHOOSECOLS: {
    method: "choosecols",
    sizeOfResultArrayMethod: "choosecolsArraySize",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_INTEGER],
    repeatLastArgs: 1,
    vectorizationForbidden: true,
  },
  TOCOL: {
    method: "tocol",
    sizeOfResultArrayMethod: "tocolArraySize",
    parameters: [WEBSHEET_FORMULA_RANGE, { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 }, { ...WEBSHEET_FORMULA_BOOLEAN, defaultValue: false }],
    vectorizationForbidden: true,
  },
  TOROW: {
    method: "torow",
    sizeOfResultArrayMethod: "torowArraySize",
    parameters: [WEBSHEET_FORMULA_RANGE, { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 }, { ...WEBSHEET_FORMULA_BOOLEAN, defaultValue: false }],
    vectorizationForbidden: true,
  },
  WRAPROWS: {
    method: "wraprows",
    sizeOfResultArrayMethod: "wraprowsArraySize",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_INTEGER, { ...WEBSHEET_FORMULA_ANY, defaultValue: webSheetFormulaError(ErrorType.NA) }],
    vectorizationForbidden: true,
  },
  WRAPCOLS: {
    method: "wrapcols",
    sizeOfResultArrayMethod: "wrapcolsArraySize",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_INTEGER, { ...WEBSHEET_FORMULA_ANY, defaultValue: webSheetFormulaError(ErrorType.NA) }],
    vectorizationForbidden: true,
  },
  EXPAND: {
    method: "expand",
    sizeOfResultArrayMethod: "expandArraySize",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_INTEGER, { ...WEBSHEET_FORMULA_INTEGER, optionalArg: true }, { ...WEBSHEET_FORMULA_ANY, defaultValue: webSheetFormulaError(ErrorType.NA) }],
    vectorizationForbidden: true,
  },
  CONVERT: {
    method: "convert",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_STRING, WEBSHEET_FORMULA_STRING],
  },
  NETWORKDAYS: {
    method: "networkdays",
    parameters: [
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL },
    ],
  },
  "NETWORKDAYS.INTL": {
    method: "networkdaysIntl",
    parameters: [
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      { ...WEBSHEET_FORMULA_ANY, defaultValue: 1 },
      { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL },
    ],
  },
  WORKDAY: {
    method: "workday",
    parameters: [
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_INTEGER,
      { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL },
    ],
  },
  "WORKDAY.INTL": {
    method: "workdayIntl",
    parameters: [
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_INTEGER,
      { ...WEBSHEET_FORMULA_ANY, defaultValue: 1 },
      { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL },
    ],
  },
  DATEVALUE: {
    method: "datevalue",
    parameters: [WEBSHEET_FORMULA_STRING],
  },
  TIMEVALUE: {
    method: "timevalue",
    parameters: [WEBSHEET_FORMULA_STRING],
  },
  HOUR: {
    method: "hour",
    parameters: [WEBSHEET_FORMULA_SCALAR],
  },
  MINUTE: {
    method: "minute",
    parameters: [WEBSHEET_FORMULA_SCALAR],
  },
  SECOND: {
    method: "second",
    parameters: [WEBSHEET_FORMULA_SCALAR],
  },
  WEEKDAY: {
    method: "weekday",
    parameters: [WEBSHEET_FORMULA_SCALAR, { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 1 }],
  },
  WEEKNUM: {
    method: "weeknum",
    parameters: [WEBSHEET_FORMULA_SCALAR, { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 1 }],
  },
  ISOWEEKNUM: {
    method: "isoweeknum",
    parameters: [WEBSHEET_FORMULA_SCALAR],
  },
  EDATE: {
    method: "edate",
    parameters: [WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_INTEGER],
  },
  EOMONTH: {
    method: "eomonth",
    parameters: [WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_INTEGER],
  },
  DAYS: {
    method: "days",
    parameters: [WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_SCALAR],
  },
  DAYS360: {
    method: "days360",
    parameters: [
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      { ...WEBSHEET_FORMULA_BOOLEAN, defaultValue: false },
    ],
  },
  YEARFRAC: {
    method: "yearfrac",
    parameters: [
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
    ],
  },
  DATEDIF: {
    method: "datedif",
    parameters: [WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_STRING],
  },
  "CEILING.MATH": {
    method: "ceilingMath",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 1 },
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 0 },
    ],
  },
  "FLOOR.MATH": {
    method: "floorMath",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 1 },
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 0 },
    ],
  },
  MROUND: {
    method: "mround",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  "CEILING.PRECISE": {
    method: "ceilingPrecise",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 1 },
    ],
  },
  "FLOOR.PRECISE": {
    method: "floorPrecise",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 1 },
    ],
  },
  QUOTIENT: {
    method: "quotient",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  SIGN: {
    method: "sign",
    parameters: [WEBSHEET_FORMULA_NUMBER],
  },
  ROMAN: {
    method: "roman",
    parameters: [WEBSHEET_FORMULA_NUMBER, { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 }],
  },
  ARABIC: {
    method: "arabic",
    parameters: [WEBSHEET_FORMULA_STRING],
  },
  BASE: {
    method: "baseFunction",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_INTEGER, { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 }],
  },
  DECIMAL: {
    method: "decimal",
    parameters: [WEBSHEET_FORMULA_STRING, WEBSHEET_FORMULA_INTEGER],
  },
  DEC2BIN: {
    method: "dec2bin",
    parameters: [WEBSHEET_FORMULA_NUMBER, { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL }],
  },
  DEC2OCT: {
    method: "dec2oct",
    parameters: [WEBSHEET_FORMULA_NUMBER, { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL }],
  },
  DEC2HEX: {
    method: "dec2hex",
    parameters: [WEBSHEET_FORMULA_NUMBER, { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL }],
  },
  BIN2DEC: {
    method: "bin2dec",
    parameters: [WEBSHEET_FORMULA_STRING],
  },
  BIN2OCT: {
    method: "bin2oct",
    parameters: [WEBSHEET_FORMULA_STRING, { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL }],
  },
  BIN2HEX: {
    method: "bin2hex",
    parameters: [WEBSHEET_FORMULA_STRING, { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL }],
  },
  OCT2BIN: {
    method: "oct2bin",
    parameters: [WEBSHEET_FORMULA_STRING, { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL }],
  },
  OCT2DEC: {
    method: "oct2dec",
    parameters: [WEBSHEET_FORMULA_STRING],
  },
  OCT2HEX: {
    method: "oct2hex",
    parameters: [WEBSHEET_FORMULA_STRING, { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL }],
  },
  HEX2BIN: {
    method: "hex2bin",
    parameters: [WEBSHEET_FORMULA_STRING, { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL }],
  },
  HEX2OCT: {
    method: "hex2oct",
    parameters: [WEBSHEET_FORMULA_STRING, { ...WEBSHEET_FORMULA_ANY, defaultValue: WEBSHEET_FORMULA_NA_SENTINEL }],
  },
  HEX2DEC: {
    method: "hex2dec",
    parameters: [WEBSHEET_FORMULA_STRING],
  },
  BITAND: {
    method: "bitand",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  BITOR: {
    method: "bitor",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  BITXOR: {
    method: "bitxor",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  BITLSHIFT: {
    method: "bitlshift",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  BITRSHIFT: {
    method: "bitrshift",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  DELTA: {
    method: "delta",
    parameters: [WEBSHEET_FORMULA_NUMBER, { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 0 }],
  },
  GESTEP: {
    method: "gestep",
    parameters: [WEBSHEET_FORMULA_NUMBER, { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 0 }],
  },
  UNICHAR: {
    method: "unichar",
    parameters: [WEBSHEET_FORMULA_INTEGER],
  },
  UNICODE: {
    method: "unicode",
    parameters: [WEBSHEET_FORMULA_STRING],
  },
  PMT: {
    method: "pmt",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 0 },
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
    ],
  },
  PV: {
    method: "pv",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 0 },
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
    ],
  },
  FV: {
    method: "fv",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 0 },
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
    ],
  },
  NPER: {
    method: "nper",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 0 },
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
    ],
  },
  RATE: {
    method: "rate",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 0 },
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 0.1 },
    ],
  },
  NPV: {
    method: "npv",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  FVSCHEDULE: {
    method: "fvSchedule",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_RANGE],
    vectorizationForbidden: true,
  },
  XNPV: {
    method: "xnpv",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_RANGE],
    vectorizationForbidden: true,
  },
  XIRR: {
    method: "xirr",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_RANGE, { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 0.1 }],
    vectorizationForbidden: true,
  },
  IRR: {
    method: "irr",
    parameters: [WEBSHEET_FORMULA_ANY, { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 0.1 }],
  },
  MIRR: {
    method: "mirr",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  DB: {
    method: "db",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 12 },
    ],
  },
  DDB: {
    method: "ddb",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 2 },
    ],
  },
  VDB: {
    method: "vdb",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 2 },
      { ...WEBSHEET_FORMULA_BOOLEAN, defaultValue: false },
    ],
  },
  AMORLINC: {
    method: "amorlinc",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
    ],
  },
  AMORDEGRC: {
    method: "amordegrc",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
    ],
  },
  SLN: {
    method: "sln",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  SYD: {
    method: "syd",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  DOLLARDE: {
    method: "dollarde",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  DOLLARFR: {
    method: "dollarfr",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  RRI: {
    method: "rri",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  PDURATION: {
    method: "pduration",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  ISPMT: {
    method: "ispmt",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  TBILLEQ: {
    method: "tbillEq",
    parameters: [WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_NUMBER],
  },
  TBILLPRICE: {
    method: "tbillPrice",
    parameters: [WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_NUMBER],
  },
  TBILLYIELD: {
    method: "tbillYield",
    parameters: [WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_NUMBER],
  },
  DISC: {
    method: "disc",
    parameters: [
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
    ],
  },
  PRICEDISC: {
    method: "priceDisc",
    parameters: [
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
    ],
  },
  YIELDDISC: {
    method: "yieldDisc",
    parameters: [
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
    ],
  },
  INTRATE: {
    method: "intrate",
    parameters: [
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
    ],
  },
  RECEIVED: {
    method: "received",
    parameters: [
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
    ],
  },
  ACCRINT: {
    method: "accrint",
    parameters: [
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 1000 },
      WEBSHEET_FORMULA_INTEGER,
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
      { ...WEBSHEET_FORMULA_BOOLEAN, defaultValue: true },
    ],
  },
  ACCRINTM: {
    method: "accrintM",
    parameters: [
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 1000 },
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
    ],
  },
  PRICEMAT: {
    method: "priceMat",
    parameters: [
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
    ],
  },
  YIELDMAT: {
    method: "yieldMat",
    parameters: [
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
    ],
  },
  COUPDAYBS: {
    method: "coupDayBs",
    parameters: [WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_INTEGER, { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 }],
  },
  COUPDAYS: {
    method: "coupDays",
    parameters: [WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_INTEGER, { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 }],
  },
  COUPDAYSNC: {
    method: "coupDaysNc",
    parameters: [WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_INTEGER, { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 }],
  },
  COUPNCD: {
    method: "coupNcd",
    parameters: [WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_INTEGER, { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 }],
  },
  COUPPCD: {
    method: "coupPcd",
    parameters: [WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_INTEGER, { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 }],
  },
  COUPNUM: {
    method: "coupNum",
    parameters: [WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_INTEGER, { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 }],
  },
  PRICE: {
    method: "price",
    parameters: [
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_INTEGER,
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
    ],
  },
  YIELD: {
    method: "yieldSecurity",
    parameters: [
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_INTEGER,
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
    ],
  },
  DURATION: {
    method: "duration",
    parameters: [
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_INTEGER,
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
    ],
  },
  MDURATION: {
    method: "mduration",
    parameters: [
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_SCALAR,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_INTEGER,
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
    ],
  },
  IPMT: {
    method: "ipmt",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 0 },
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
    ],
  },
  PPMT: {
    method: "ppmt",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      { ...WEBSHEET_FORMULA_NUMBER, defaultValue: 0 },
      { ...WEBSHEET_FORMULA_INTEGER, defaultValue: 0 },
    ],
  },
  CUMIPMT: {
    method: "cumipmt",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_INTEGER,
    ],
  },
  CUMPRINC: {
    method: "cumprinc",
    parameters: [
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_NUMBER,
      WEBSHEET_FORMULA_INTEGER,
    ],
  },
  EFFECT: {
    method: "effect",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  NOMINAL: {
    method: "nominal",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  COMPLEX: {
    method: "complexFunction",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER, { ...WEBSHEET_FORMULA_STRING, defaultValue: "i" }],
  },
  IMABS: {
    method: "imabs",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  IMREAL: {
    method: "imreal",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  IMAGINARY: {
    method: "imaginary",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  IMSUM: {
    method: "imsum",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  IMPRODUCT: {
    method: "improduct",
    parameters: [WEBSHEET_FORMULA_ANY],
    repeatLastArgs: 1,
  },
  IMSUB: {
    method: "imsub",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY],
  },
  IMDIV: {
    method: "imdiv",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_ANY],
  },
  IMCONJUGATE: {
    method: "imconjugate",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  IMARGUMENT: {
    method: "imargument",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  IMEXP: {
    method: "imexp",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  IMLN: {
    method: "imln",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  IMLOG10: {
    method: "imlog10",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  IMLOG2: {
    method: "imlog2",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  IMPOWER: {
    method: "impowerFunction",
    parameters: [WEBSHEET_FORMULA_ANY, WEBSHEET_FORMULA_NUMBER],
  },
  IMSQRT: {
    method: "imsqrt",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  IMSIN: {
    method: "imsin",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  IMCOS: {
    method: "imcos",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  IMTAN: {
    method: "imtan",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  IMSEC: {
    method: "imsec",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  IMCSC: {
    method: "imcsc",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  IMCOT: {
    method: "imcot",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  IMSINH: {
    method: "imsinh",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  IMCOSH: {
    method: "imcosh",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  IMSECH: {
    method: "imsech",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  IMCSCH: {
    method: "imcsch",
    parameters: [WEBSHEET_FORMULA_ANY],
  },
  BESSELI: {
    method: "besseli",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  BESSELJ: {
    method: "besselj",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  BESSELK: {
    method: "besselk",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  BESSELY: {
    method: "bessely",
    parameters: [WEBSHEET_FORMULA_NUMBER, WEBSHEET_FORMULA_NUMBER],
  },
  DSUM: {
    method: "dsum",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_RANGE],
  },
  DCOUNT: {
    method: "dcount",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_RANGE],
  },
  DCOUNTA: {
    method: "dcounta",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_RANGE],
  },
  DAVERAGE: {
    method: "daverage",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_RANGE],
  },
  DMAX: {
    method: "dmax",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_RANGE],
  },
  DMIN: {
    method: "dmin",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_RANGE],
  },
  DPRODUCT: {
    method: "dproduct",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_RANGE],
  },
  DSTDEV: {
    method: "dstdev",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_RANGE],
  },
  DSTDEVP: {
    method: "dstdevp",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_RANGE],
  },
  DVAR: {
    method: "dvar",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_RANGE],
  },
  DVARP: {
    method: "dvarp",
    parameters: [WEBSHEET_FORMULA_RANGE, WEBSHEET_FORMULA_SCALAR, WEBSHEET_FORMULA_RANGE],
  },
};



const WEBSHEET_FORMULA_TRANSLATIONS = {
  enGB: Object.fromEntries(Object.keys(WebSheetFormulaPlugin.implementedFunctions).map((name) => [name, name])),
};



registerWebSheetFormulaPlugin();


const MACRO_SECURITY_LEVELS = new Set(["disableAll", "disableWithNotification", "signedOnly", "enableAll"]);


const MACRO_SECURITY_OPTIONS = [
  {
    value: "disableAll",
    label: "警告を表示せずにすべての VBA マクロを無効にする",
    description: "信頼できないブックを開いたときにマクロの通知を表示しません。",
  },
  {
    value: "disableWithNotification",
    label: "警告を表示してすべての VBA マクロを無効にする",
    description: "Excel と同様に、マクロを含むブックで通知を表示します。",
  },
  {
    value: "signedOnly",
    label: "デジタル署名されたマクロを除き、すべての VBA マクロを無効にする",
    description: "信頼された発行元の署名付きマクロだけを許可する設定です。",
  },
  {
    value: "enableAll",
    label: "すべての VBA マクロを有効にする (推奨しません。危険なコードが実行される可能性があります)",
    description: "互換性確認用です。WebSheet では対応済みの安全な VBA サブセットだけを実行します。",
  },
];


const MACRO_SECURITY_NAV_ITEMS = [
  "信頼できる発行元",
  "信頼できる場所",
  "信頼済みドキュメント",
  "アドイン",
  "ActiveX の設定",
  "マクロの設定",
  "メッセージ バー",
  "外部コンテンツ",
  "ファイル制限機能の設定",
  "プライバシー オプション",
];


const SHEET_TAB_COLORS = [
  { label: "色なし", value: "" },
  { label: "赤", value: "#c00000" },
  { label: "橙", value: "#ed7d31" },
  { label: "黄", value: "#ffc000" },
  { label: "緑", value: "#70ad47" },
  { label: "青", value: "#4472c4" },
  { label: "紫", value: "#7030a0" },
  { label: "灰", value: "#808080" },
];


const CHART_COLORS = ["#4472c4", "#ed7d31", "#70ad47", "#a5a5a5", "#ffc000", "#5b9bd5", "#7030a0", "#00b0f0"];


const RIBBON_COLOR_TINTS = [0.82, 0.62, 0.42, -0.25, -0.5];


const RIBBON_STANDARD_COLORS = ["#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050", "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0"];


const DEFAULT_THEME_COLORS = [
  "FFFFFF",
  "000000",
  "EEECE1",
  "1F497D",
  "4F81BD",
  "C0504D",
  "9BBB59",
  "8064A2",
  "4BACC6",
  "F79646",
  "0000FF",
  "800080",
];


const INSERT_SYMBOL_CHOICES = ["Ω", "Σ", "π", "μ", "±", "×", "÷", "≠", "≈", "≤", "≥", "℃", "㎡", "①", "②", "③", "→", "←", "↑", "↓", "■", "□", "●", "○"];


const DRAW_INK_TOOLS = {
  pen: { label: "ペン", color: "#111827", width: 2.5, opacity: 1 },
  pencil: { label: "鉛筆", color: "#374151", width: 1.5, opacity: 0.9 },
  highlighter: { label: "蛍光ペン", color: "#facc15", width: 10, opacity: 0.35 },
};


const DRAW_INK_ERASERS = {
  stroke: { label: "ストローク消しゴム", menuLabel: "ストローク消しゴム(S)", shortLabel: "消しゴム", radius: 8, mode: "stroke" },
  point: { label: "ポイント消しゴム", menuLabel: "ポイント消しゴム(P)", shortLabel: "ポイント", radius: 10, mode: "point" },
  small: { label: "小さい消しゴム", menuLabel: "小さい消しゴム(K)", shortLabel: "小", radius: 7, mode: "point" },
  medium: { label: "中くらいの消しゴム", menuLabel: "中くらいの消しゴム(M)", shortLabel: "中", radius: 14, mode: "point" },
  large: { label: "大きい消しゴム", menuLabel: "大きい消しゴム(L)", shortLabel: "大", radius: 24, mode: "point" },
};


const DRAW_COLOR_PRESETS = {
  black: "#111827",
  blue: "#2563eb",
  red: "#dc2626",
  yellow: "#facc15",
};


const PAGE_LAYOUT_THEMES = [
  { value: "office", label: "Office", colors: ["FFFFFF", "000000", "EEECE1", "1F497D", "4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646", "0000FF", "800080"], fonts: { heading: "Calibri Light", body: "Calibri" } },
  { value: "office2013", label: "Office 2013-2022", colors: ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72"], fonts: { heading: "Aptos Display", body: "Aptos" } },
  { value: "ion", label: "イオン", colors: ["#2f75b5", "#ed7d31", "#70ad47", "#ffc000", "#5b9bd5", "#a64d79"], fonts: { heading: "Aptos Display", body: "Aptos" } },
  { value: "facet", label: "ファセット", colors: ["#4472c4", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47", "#ed7d31"], fonts: { heading: "Yu Gothic", body: "Yu Gothic UI" } },
  { value: "gallery", label: "ギャラリー", colors: ["#2B579A", "#D24726", "#008A00", "#5133AB", "#AC193D", "#2672EC"], fonts: { heading: "Century Gothic", body: "Aptos" } },
  { value: "integral", label: "インテグラル", colors: ["#4F81BD", "#C0504D", "#9BBB59", "#8064A2", "#4BACC6", "#F79646"], fonts: { heading: "Georgia", body: "Calibri" } },
  { value: "organic", label: "オーガニック", colors: ["#546A3A", "#A4A700", "#7E8F6A", "#C0504D", "#9BBB59", "#F79646"], fonts: { heading: "Candara", body: "Calibri" } },
  { value: "parcel", label: "パーセル", colors: ["#c00000", "#7030a0", "#00b0f0", "#92d050", "#ffc000", "#ed7d31"], fonts: { heading: "MS Pゴシック", body: "MS Pゴシック" } },
  { value: "retrospect", label: "レトロスペクト", colors: ["#E84C22", "#F6C141", "#2A8C83", "#0D5C75", "#7A3D9A", "#C03A2B"], fonts: { heading: "Book Antiqua", body: "Century Gothic" } },
  { value: "slice", label: "スライス", colors: ["#5B9BD5", "#ED7D31", "#A5A5A5", "#FFC000", "#4472C4", "#70AD47"], fonts: { heading: "Trebuchet MS", body: "Aptos" } },
  { value: "wisp", label: "ウィスプ", colors: ["#7C9BA6", "#A67C52", "#8FAADC", "#C9A45C", "#9E7BB5", "#6AA84F"], fonts: { heading: "Gill Sans MT", body: "Calibri" } },
];


const PAGE_LAYOUT_EFFECT_PRESETS = [
  { value: "office", label: "Office", description: "標準の影と控えめな立体感", shape: { shadow: "outer1", glow: null, softEdges: 0, reflection: null } },
  { value: "office2013", label: "Office 2013-2022", description: "フラットに近い軽い影", shape: { shadow: "outer2", glow: null, softEdges: 0, reflection: null } },
  { value: "ion", label: "イオン", description: "青い光彩と弱い影", shape: { shadow: { dx: 1.5, dy: 2, blur: 3.5, color: "rgba(31,78,121,0.28)" }, glow: "blue", softEdges: 0, reflection: null } },
  { value: "facet", label: "ファセット", description: "シャープな影と少し強い輪郭感", shape: { shadow: { dx: 2.5, dy: 2.5, blur: 2.2, color: "rgba(0,0,0,0.26)" }, glow: null, softEdges: 0, reflection: null } },
  { value: "gallery", label: "ギャラリー", description: "反射を含むギャラリー風", shape: { shadow: "outer1", glow: null, softEdges: 0, reflection: "tight" } },
  { value: "integral", label: "インテグラル", description: "強めの影と柔らかい端", shape: { shadow: { dx: 3, dy: 3, blur: 5, color: "rgba(0,0,0,0.32)" }, glow: null, softEdges: 0.5, reflection: null } },
  { value: "organic", label: "オーガニック", description: "緑系の光彩と柔らかい端", shape: { shadow: "outer2", glow: { radius: 4, color: "#70ad47" }, softEdges: 0.6, reflection: null } },
  { value: "parcel", label: "パーセル", description: "はっきりした影と軽い反射", shape: { shadow: { dx: 2, dy: 3, blur: 4, color: "rgba(0,0,0,0.34)" }, glow: null, softEdges: 0, reflection: "tight" } },
  { value: "retrospect", label: "レトロスペクト", description: "暖色の光彩と影", shape: { shadow: { dx: 3, dy: 2, blur: 4, color: "rgba(132,60,12,0.32)" }, glow: { radius: 5, color: "#ed7d31" }, softEdges: 0, reflection: null } },
  { value: "slice", label: "スライス", description: "クリーンな影と薄いぼかし", shape: { shadow: "outer1", glow: null, softEdges: 0.4, reflection: null } },
  { value: "wisp", label: "ウィスプ", description: "薄い光彩と柔らかい反射", shape: { shadow: "outer2", glow: { radius: 3, color: "#8faadc" }, softEdges: 0.5, reflection: "tight" } },
  { value: "flat", label: "フラット", description: "影や光彩を使用しない", shape: { shadow: "", glow: null, softEdges: 0, reflection: null } },
];


const EXCEL_DEFAULT_PAGE_MARGINS = Object.freeze({ top: 0.75, bottom: 0.75, left: 0.7, right: 0.7, header: 0.3, footer: 0.3 });


const PAGE_LAYOUT_MARGIN_PRESETS = {
  normal: { label: "標準", ...EXCEL_DEFAULT_PAGE_MARGINS },
  wide: { label: "広い", top: 1, bottom: 1, left: 1, right: 1, header: 0.5, footer: 0.5 },
  narrow: { label: "狭い", top: 0.25, bottom: 0.25, left: 0.25, right: 0.25, header: 0.3, footer: 0.3 },
};


const PAGE_LAYOUT_PAPER_SIZES = [
  { value: "A4", label: "A4", width: 8.27, height: 11.69 },
  { value: "A3", label: "A3", width: 11.69, height: 16.54 },
  { value: "B5", label: "B5", width: 7.17, height: 10.12 },
  { value: "Letter", label: "レター", width: 8.5, height: 11 },
];


const RIBBON_COLOR_STYLE = {
  fillColor: { cssProperty: "backgroundColor", label: "背景色", defaultColor: "#fff2cc" },
  fontColor: { cssProperty: "color", label: "文字色", defaultColor: "#000000" },
};


const RIBBON_BORDER_PRESETS = [
  { value: "none", label: "枠線なし" },
  { value: "outside", label: "外枠" },
  { value: "thick-outside", label: "外枠太罫線" },
  { value: "all", label: "格子" },
  { value: "top", label: "上罫線" },
  { value: "right", label: "右罫線" },
  { value: "bottom", label: "下罫線" },
  { value: "left", label: "左罫線" },
  { value: "thick-bottom", label: "太い下罫線" },
];


const RIBBON_UNDERLINE_PRESETS = [
  { value: "underline", label: "下線(U)", mark: "U" },
  { value: "underline double", label: "二重下線(D)", mark: "D" },
];


const DEFAULT_SHAPE_TYPE = "rect";


const CUSTOM_SHAPE_TYPE = "customPath";


const SHAPE_FIT_INSET = 0.2;


const DEFAULT_RECENT_SHAPE_TYPES = ["rect", "line", "rightBlockArrow", "oval", "calloutRect"];


const RIBBON_SHAPE_GROUPS = [
  {
    key: "lines",
    label: "線",
    shapes: [
      { type: "line", label: "線" },
      { type: "arrow", label: "矢印" },
      { type: "doubleArrow", label: "双方向矢印" },
      { type: "elbowConnector", label: "カギ線コネクタ" },
      { type: "curvedConnector", label: "曲線コネクタ" },
      { type: "freeform", label: "フリーフォーム" },
    ],
  },
  {
    key: "rectangles",
    label: "四角形",
    shapes: [
      { type: "rect", label: "四角形" },
      { type: "textBox", label: "テキスト ボックス" },
      { type: "roundRect", label: "角丸四角形" },
      { type: "snipRect", label: "切り欠き四角形" },
      { type: "roundSnipRect", label: "角丸切り欠き四角形" },
      { type: "frame", label: "フレーム" },
    ],
  },
  {
    key: "basic",
    label: "基本図形",
    shapes: [
      { type: "oval", label: "楕円" },
      { type: "triangle", label: "二等辺三角形" },
      { type: "rightTriangle", label: "直角三角形" },
      { type: "diamond", label: "ひし形" },
      { type: "pentagon", label: "五角形" },
      { type: "hexagon", label: "六角形" },
      { type: "heptagon", label: "七角形" },
      { type: "octagon", label: "八角形" },
      { type: "decagon", label: "十角形" },
      { type: "dodecagon", label: "十二角形" },
      { type: "parallelogram", label: "平行四辺形" },
      { type: "trapezoid", label: "台形" },
      { type: "cross", label: "十字形" },
      { type: "can", label: "円柱" },
      { type: "cube", label: "直方体" },
      { type: "bevel", label: "額縁" },
      { type: "donut", label: "ドーナツ" },
      { type: "noSymbol", label: "禁止" },
      { type: "blockArc", label: "ブロック弧" },
      { type: "foldedCorner", label: "メモ" },
      { type: "smileyFace", label: "スマイル" },
      { type: "heart", label: "ハート" },
      { type: "lightningBolt", label: "稲妻" },
      { type: "sun", label: "太陽" },
      { type: "moon", label: "月" },
      { type: "cloud", label: "雲" },
      { type: "arc", label: "円弧" },
      { type: "pie", label: "円グラフ" },
      { type: "chord", label: "弦" },
      { type: "teardrop", label: "涙" },
      { type: "plaque", label: "額縁" },
      { type: "corner", label: "L 字" },
      { type: "halfFrame", label: "半フレーム" },
      { type: "diagStripe", label: "斜め縞" },
      { type: "leftBracket", label: "左大かっこ" },
      { type: "rightBracket", label: "右大かっこ" },
      { type: "bracketPair", label: "大かっこ" },
      { type: "leftBrace", label: "左中かっこ" },
      { type: "rightBrace", label: "右中かっこ" },
      { type: "bracePair", label: "中かっこ" },
    ],
  },
  {
    key: "blockArrows",
    label: "ブロック矢印",
    shapes: [
      { type: "rightBlockArrow", label: "右矢印" },
      { type: "leftBlockArrow", label: "左矢印" },
      { type: "upBlockArrow", label: "上矢印" },
      { type: "downBlockArrow", label: "下矢印" },
      { type: "leftRightBlockArrow", label: "左右矢印" },
      { type: "leftRightUpArrow", label: "左右上下矢印" },
      { type: "quadArrow", label: "四方向矢印" },
      { type: "notchedRightArrow", label: "切り込み右矢印" },
      { type: "bentUpArrow", label: "曲折上矢印" },
      { type: "uturnArrow", label: "U ターン矢印" },
      { type: "curvedRightArrow", label: "右カーブ矢印" },
      { type: "leftUpArrow", label: "左上矢印" },
      { type: "stripedRightArrow", label: "縞付き右矢印" },
      { type: "homePlate", label: "ホームベース" },
      { type: "chevron", label: "山形" },
      { type: "leftRightCircularArrow", label: "左右循環矢印" },
    ],
  },
  {
    key: "equations",
    label: "数式図形",
    shapes: [
      { type: "mathPlus", label: "プラス" },
      { type: "mathMinus", label: "マイナス" },
      { type: "mathMultiply", label: "乗算" },
      { type: "mathDivide", label: "除算" },
      { type: "mathEqual", label: "等号" },
      { type: "mathNotEqual", label: "不等号" },
    ],
  },
  {
    key: "flowchart",
    label: "フローチャート",
    shapes: [
      { type: "flowProcess", label: "処理" },
      { type: "flowDecision", label: "判断" },
      { type: "flowData", label: "データ" },
      { type: "flowDocument", label: "書類" },
      { type: "flowMultiDocument", label: "複数書類" },
      { type: "flowTerminator", label: "端子" },
      { type: "flowPredefinedProcess", label: "定義済み処理" },
      { type: "flowStoredData", label: "記憶データ" },
      { type: "flowInternalStorage", label: "内部記憶" },
      { type: "flowManualInput", label: "手入力" },
      { type: "flowManualOperation", label: "手操作" },
      { type: "flowPreparation", label: "準備" },
      { type: "flowConnector", label: "結合子" },
      { type: "flowOffpageConnector", label: "他ページ結合子" },
      { type: "flowDelay", label: "遅延" },
      { type: "flowDisplay", label: "表示" },
      { type: "flowOr", label: "論理和" },
      { type: "flowSummingJunction", label: "和接合" },
    ],
  },
  {
    key: "stars",
    label: "星とリボン",
    shapes: [
      { type: "star5", label: "星5" },
      { type: "star4", label: "星4" },
      { type: "star6", label: "星6" },
      { type: "star7", label: "星7" },
      { type: "star8", label: "星8" },
      { type: "star10", label: "星10" },
      { type: "star12", label: "星12" },
      { type: "star16", label: "星16" },
      { type: "star24", label: "星24" },
      { type: "star32", label: "星32" },
      { type: "seal", label: "爆発1" },
      { type: "explosion", label: "爆発2" },
      { type: "ribbon", label: "リボン" },
      { type: "wave", label: "波" },
    ],
  },
  {
    key: "callouts",
    label: "吹き出し",
    shapes: [
      { type: "calloutRect", label: "四角形吹き出し" },
      { type: "calloutRoundRect", label: "角丸四角形吹き出し" },
      { type: "calloutOval", label: "楕円吹き出し" },
      { type: "calloutCloud", label: "雲形吹き出し" },
    ],
  },
];


const RIBBON_SHAPE_PRESET_MAP = new Map(
  RIBBON_SHAPE_GROUPS.flatMap((group) => group.shapes).map((shape) => [shape.type, shape]),
);


const DEFAULT_SHAPE_OUTLINE_WIDTH = 0.75;


const LINE_SHAPE_TYPES = new Set(["line", "arrow", "doubleArrow", "elbowConnector", "curvedConnector", "freeform"]);


const SHAPE_STYLE_PRESETS = {
  softBlue: { fillColor: "#dbeafe", outlineColor: "#4472c4", outlineWidth: 0.75, textColor: "#1f2937" },
  softGreen: { fillColor: "#e2f0d9", outlineColor: "#70ad47", outlineWidth: 0.75, textColor: "#1f2937" },
  softOrange: { fillColor: "#fce4d6", outlineColor: "#ed7d31", outlineWidth: 0.75, textColor: "#1f2937" },
  darkBlue: { fillColor: "#1f4e79", outlineColor: "#1f4e79", outlineWidth: 0.75, textColor: "#ffffff" },
};


const SHAPE_OUTLINE_WEIGHTS = [0.75, 1, 1.5, 2.25, 3, 4.5, 6];


const SHAPE_LINE_DASH_PRESETS = [
  { value: "solid", label: "実線" },
  { value: "shortDash", label: "短い破線" },
  { value: "dash", label: "破線" },
  { value: "dot", label: "点線" },
  { value: "squareDot", label: "四角点" },
  { value: "dashDot", label: "一点鎖線" },
  { value: "longDash", label: "長い破線" },
  { value: "longDashDot", label: "長い一点鎖線" },
  { value: "longDashDotDot", label: "長い二点鎖線" },
];


const SHAPE_LINE_COMPOUND_PRESETS = [
  { value: "single", label: "一重線" },
  { value: "double", label: "二重線" },
  { value: "thickThin", label: "太細線" },
  { value: "thinThick", label: "細太線" },
  { value: "triple", label: "三重線" },
];


const SHAPE_LINE_CAP_PRESETS = [
  { value: "butt", label: "フラット" },
  { value: "square", label: "四角" },
  { value: "round", label: "丸" },
];


const SHAPE_LINE_JOIN_PRESETS = [
  { value: "round", label: "丸" },
  { value: "bevel", label: "面取り" },
  { value: "miter", label: "角" },
];


const SHAPE_ARROW_PRESETS = [
  { value: "none", label: "矢印なし", start: false, end: false },
  { value: "end", label: "終点矢印", start: false, end: true },
  { value: "start", label: "始点矢印", start: true, end: false },
  { value: "both", label: "両端矢印", start: true, end: true },
];


const SHAPE_ARROW_TYPE_PRESETS = [
  { value: "none", label: "なし" },
  { value: "triangle", label: "三角矢印" },
  { value: "open", label: "開いた矢印" },
  { value: "stealth", label: "ステルス矢印" },
  { value: "diamond", label: "ひし形矢印" },
  { value: "oval", label: "楕円矢印" },
];


const SHAPE_ARROW_SIZE_PRESETS = [
  { value: "small", label: "小" },
  { value: "medium", label: "中" },
  { value: "large", label: "大" },
];


const SHAPE_FILL_PATTERN_PRESETS = [
  { value: "", label: "パターンなし" },
  { value: "pct5", label: "5%" },
  { value: "pct10", label: "10%" },
  { value: "pct20", label: "20%" },
  { value: "pct25", label: "25%" },
  { value: "pct50", label: "50%" },
  { value: "ltHorz", label: "横線" },
  { value: "ltVert", label: "縦線" },
  { value: "ltDnDiag", label: "右下がり斜線" },
  { value: "ltUpDiag", label: "右上がり斜線" },
  { value: "diagCross", label: "斜め格子" },
  { value: "smGrid", label: "小さい格子" },
];


const SHAPE_GLOW_PRESETS = [
  { value: "", label: "光彩なし" },
  { value: "blue", label: "青 5 pt", radius: 5, color: "#5b9bd5" },
  { value: "green", label: "緑 5 pt", radius: 5, color: "#70ad47" },
  { value: "orange", label: "オレンジ 5 pt", radius: 5, color: "#ed7d31" },
  { value: "purple", label: "紫 8 pt", radius: 8, color: "#8064a2" },
];


const SHAPE_SOFT_EDGE_PRESETS = [
  { value: 0, label: "なし" },
  { value: 1, label: "1 pt" },
  { value: 2.5, label: "2.5 pt" },
  { value: 5, label: "5 pt" },
  { value: 10, label: "10 pt" },
];


const SHAPE_REFLECTION_PRESETS = [
  { value: "", label: "反射なし" },
  { value: "tight", label: "密接反射" },
  { value: "half", label: "半反射" },
  { value: "full", label: "全反射" },
];


const STATUS_BAR_STORAGE_KEY = "websheet.statusBarItems";


const INTERNAL_CLIPBOARD_STORAGE_KEY = "websheet.internalClipboard.v1";


const RIBBON_LAYOUT_STORAGE_KEY = "websheet.ribbonLayout.v1";


const RIBBON_VISIBILITY_STORAGE_KEY = "websheet.ribbonVisibility.v1";


const RIBBON_LAYOUTS = new Set(["classic", "single"]);


const RIBBON_VISIBILITY_MODES = new Set(["always", "auto"]);


const CELL_EDIT_TRAILING_LINE_SENTINEL = "\u200b";


const STATUS_BAR_ITEMS = [
  { key: "average", label: "平均", type: "stat" },
  { key: "count", label: "データの個数", type: "stat" },
  { key: "numericCount", label: "数値の個数", type: "stat" },
  { key: "min", label: "最小値", type: "stat" },
  { key: "max", label: "最大値", type: "stat" },
  { key: "sum", label: "合計", type: "stat" },
  { type: "separator" },
  { key: "viewShortcuts", label: "表示ショートカット", type: "control" },
  { key: "zoomSlider", label: "ズーム スライダー", type: "control" },
  { key: "zoom", label: "ズーム", type: "control" },
];


const DEFAULT_STATUS_BAR_ITEMS = Object.freeze({
  average: true,
  count: true,
  numericCount: false,
  min: false,
  max: false,
  sum: true,
  viewShortcuts: true,
  zoomSlider: true,
  zoom: true,
});


const EXCEL_MAX_ROWS = 1048576;


const EXCEL_MAX_COLS = 16384;


const ZOOM_DIALOG_PRESETS = [
  { value: 200, label: "200%(0)" },
  { value: 100, label: "100%(1)" },
  { value: 75, label: "75%(7)" },
  { value: 50, label: "50%(5)" },
  { value: 25, label: "25%(2)" },
];


const LINK_DIALOG_KINDS = [
  { kind: "web", label: "ファイル、Web ページ", fieldLabel: "アドレス(E):", placeholder: "https://example.com" },
  { kind: "document", label: "このドキュメント内", fieldLabel: "セル参照を入力してください(E):", placeholder: "A1" },
  { kind: "email", label: "電子メール アドレス", fieldLabel: "電子メール アドレス(E):", placeholder: "name@example.com" },
];


const CELL_BORDER_SIDES = ["top", "right", "bottom", "left"];


const DEFAULT_CELL_CSS_KEYS = [
  "fontFamily",
  "fontSize",
  "color",
  "fontWeight",
  "fontStyle",
  "textDecoration",
  "textAlign",
  "verticalAlign",
  "alignItems",
  "textRotation",
];


const INFERRED_DEFAULT_CELL_CSS_KEYS = new Set(["fontFamily", "fontSize"]);


const LOCALIZED_DEFAULT_CELL_CSS_KEYS = new Set([
  "color",
  "fontWeight",
  "fontStyle",
  "textDecoration",
  "textAlign",
  "alignItems",
  "textRotation",
]);


const EMPTY_CELL_CSS = Object.freeze({});


const EMPTY_CELL_BORDERS = Object.freeze({});


const cellStyleInternPools = {
  css: new Map(),
  borderSide: new Map(),
  borders: new Map(),
};


const formulaSpillAnchorCacheBySheet = new WeakMap();


const sheetUsedRangeCacheBySheet = new WeakMap();


const sheetStyleRangeCacheBySheet = new WeakMap();


const sheetMergeMapsCacheByMerges = new WeakMap();


const sheetMinimumExtentCacheBySheet = new WeakMap();


const sheetBorderCellCacheBySheet = new WeakMap();


const sheetPersistentNoteCellCacheBySheet = new WeakMap();



const state = {
  model: null,
  isStandalone: false,
  hf: null,
  activeSheetIndex: 0,
  selectedSheetIndexes: [],
  sheetContextMenu: null,
  sheetNavigatorMenu: null,
  sheetTabRename: null,
  sheetTabDrag: null,
  sheetTabKeyboardActivation: null,
  suppressNextSheetTabClick: false,
  cellContextMenu: null,
  selected: null,
  selectedImage: null,
  selectedImages: [],
  selectionRange: null,
  selectionRanges: [],
  selectionAnchor: null,
  selectionStart: null,
  selectionDrag: null,
  selectionAutoScroll: null,
  objectSelectionMode: false,
  objectSelectionDrag: null,
  fillDrag: null,
  autoFillModifierActive: false,
  autoFillOptions: null,
  suppressAutoFillContextMenuUntil: 0,
  pageBreakDrag: null,
  printAreaDrag: null,
  isSelecting: false,
  headerResize: null,
  headerResizeFrame: null,
  formatDialog: null,
  headerSizeDialog: null,
  homeCommandMenu: null,
  viewCommandMenu: null,
  splitPaneDrag: null,
  findDialog: null,
  lastFind: null,
  dataDialog: null,
  autoFilterMenu: null,
  nameDialog: null,
  linkDialog: null,
  objectLinkDialog: null,
  translationPane: null,
  symbolDialog: null,
  suppressNextCellClick: false,
  editAutoScroll: null,
  editAutoScrollFrame: null,
  cellEditSelection: null,
  cellEditDirectTyping: null,
  selectionKeyboardSink: null,
  cellEditComposing: false,
  formulaBarComposing: false,
  formulaEdit: null,
  nameBoxMenu: null,
  showFormulas: false,
  formulaEngineFallback: false,
  formulaAutocomplete: null,
  suppressFormulaAutocompleteInput: false,
  formulaAutocompleteAccepted: null,
  formulaReferenceRange: null,
  formulaReferenceDrag: null,
  formulaReferenceSheetNavigation: null,
  syncingScroll: false,
  restoringViewportScroll: false,
  lazyScrollLocks: { rows: false, cols: false },
  viewportRenderFrame: null,
  viewportTrimTimer: null,
  printSurfaceCleanupTimer: null,
  deferredViewportRender: false,
  deferredSelectionUiSyncTimer: null,
  renderDisplayCache: null,
  renderFormulaSpillCache: null,
  undoStack: [],
  redoStack: [],
  isRestoringHistory: false,
  cwsAgentPreview: null,
  internalClipboard: null,
  objectClipboard: null,
  consumedCutObjectClipboardLabel: null,
  pendingPasteFallback: null,
  pendingObjectPasteFallback: null,
  suppressBeforeUnloadWarning: false,
  formatPainter: null,
  ribbonColorPalette: null,
  ribbonBorderMenu: null,
  ribbonUnderlineMenu: null,
  ribbonCollapsed: false,
  ribbonLayout: "classic",
  ribbonVisibility: "always",
  ribbonAutoOpen: false,
  ribbonShapeMenu: null,
  ribbonShapeFormatMenu: null,
  ribbonInkEraserMenu: null,
  quickAnalysisPopup: null,
  messageBoxHandlers: null,
  dataCommandMenu: null,
  dataValidationMenu: null,
  validationDropdownPopup: null,
  validationInvalidMarks: {},
  whatIfMenu: null,
  queryPane: null,
  queryPaneContextMenu: null,
  queryCredentials: new Map(),
  queryRefreshTimers: [],
  activeQueryRefreshes: [],
  recentFormulaFunctionNames: [],
  shapeFormatPane: null,
  shapeSelectionPane: null,
  shapeSelectionPaneDrag: null,
  shapePointEditIds: [],
  lineConnectionPreview: null,
  currentShapeType: DEFAULT_SHAPE_TYPE,
  recentShapeTypes: [...DEFAULT_RECENT_SHAPE_TYPES],
  pendingShapeType: null,
  shapeDrawing: null,
  currentInkTool: { ...DRAW_INK_TOOLS.pen },
  currentInkColor: DRAW_INK_TOOLS.pen.color,
  currentInkEraser: "stroke",
  pendingInkTool: null,
  inkDrawing: null,
  selectedInkStrokeIds: [],
  imageTransform: null,
  pictureCrop: null,
  pictureCropDrag: null,
  shapeTextEditor: null,
  zoomDialog: null,
  pendingSheetViewScroll: null,
  workbookLoadVersion: 0,
  openedFileSignature: null,
  persistenceTimer: null,
  persistenceIdleCallback: null,
  persistencePromise: null,
  persistenceLoading: false,
  standalonePersistenceScope: null,
  formulaBarLines: FORMULA_BAR_MIN_LINES,
  formulaBarExpandedLines: FORMULA_BAR_DEFAULT_EXPANDED_LINES,
  formulaBarResize: null,
  cellStructureDialog: null,
  annotationEditor: null,
  commentPopupCloseTimer: null,
  notePopupCloseTimer: null,
  statusBarItems: { ...DEFAULT_STATUS_BAR_ITEMS },
  statusBarContextMenu: false,
  currentSaveType: "",
  currentSaveFileHandle: null,
  currentSaveDirectoryHandle: null,
  currentSaveFileName: "",
  workbookRevision: 0,
  workbookSavedRevision: 0,
  serverVersionInfo: null,
  serverVersionStatus: "idle",
  versionMismatch: false,
  helpDialog: null,
  versionDialog: null,
  dialogChromeObserver: null,
  dialogDrag: null,
};



const icons = {
  AlignCenter,
  AlignHorizontalJustifyCenter,
  AlignHorizontalJustifyEnd,
  AlignHorizontalJustifyStart,
  AlignJustify,
  AlignLeft,
  AlignRight,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  ArrowDownNarrowWide,
  ArrowDownToLine,
  ArrowUpDown,
  BadgePlus,
  BarChart3,
  BookOpen,
  Box,
  CalendarClock,
  CaseSensitive,
  ChartColumn,
  ChartColumnBig,
  ChartLine,
  ChartNoAxesColumn,
  ChartPie,
  Check,
  ChevronDown,
  ChevronUp,
  ChevronsDown,
  ChevronsUp,
  ClipboardPaste,
  Clock3,
  CircleHelp,
  Columns2,
  Columns3,
  Copy,
  Crop,
  Database,
  Delete,
  Download,
  Eraser,
  Eye,
  EyeOff,
  FilePlus,
  FileSpreadsheet,
  FileText,
  FileUp,
  Filter,
  FolderOpen,
  FunctionSquare,
  GitBranch,
  Globe,
  Grid2X2,
  Grid2x2: Grid2X2,
  Hash,
  Image: ImageIcon,
  Info,
  Italic,
  KeyRound,
  Link: LinkIcon,
  LogOut,
  ListChecks,
  ListFilter,
  MailCheck,
  Map: MapIcon,
  Maximize,
  MessageSquareText,
  MessageSquareX,
  MousePointer2,
  Minus,
  MonitorUp,
  MoveRight,
  Omega,
  Paintbrush,
  PaintBucket,
  Palette,
  PanelTop,
  PenLine,
  Pencil,
  Highlighter,
  Lasso,
  Plug,
  Printer,
  Puzzle,
  RefreshCw,
  Repeat,
  RotateCcw,
  RotateCw,
  Ruler,
  Save,
  SaveAll,
  Scissors,
  ScatterChart,
  Search,
  Send,
  Settings,
  Shapes,
  Signature,
  Sigma,
  Square,
  SquareDashed,
  SquareFunction,
  Sparkles,
  Table2,
  TableCellsMerge,
  TableCellsSplit,
  TableProperties,
  TextCursorInput,
  Trash2,
  Type,
  Underline,
  UnlockKeyhole,
  Unlink,
  Upload,
  UploadCloud,
  WrapText,
  X,
  PlaySquare,
  ZoomIn,
};



const $fileInput = $("#fileInput");


const $pageBackgroundInput = $("#pageBackgroundInput");


const $bootState = $("#bootState");


const $emptyState = $("#emptyState");


const $sheetWorkspace = $("#sheetWorkspace");


const $dropZone = $("#dropZone");


const $sheetHost = $("#sheetHost");


const $sheetTabs = $("#sheetTabs");


const $sheetFooter = $("#sheetFooter");


const $sheetFooterSplitter = $("#sheetFooterSplitter");


const $horizontalScrollbar = $("#horizontalScrollbar");


const $horizontalScrollSpacer = $("#horizontalScrollSpacer");


const $addSheetButton = $("#addSheetButton");


const $sheetTabScrollLeft = $("#sheetTabScrollLeft");


const $sheetTabScrollRight = $("#sheetTabScrollRight");


const $zoomOutButton = $("#zoomOutButton");


const $zoomInButton = $("#zoomInButton");


const $zoomSlider = $("#zoomSlider");


const $zoomLabel = $("#zoomLabel");


const $selectionStatus = $("#selectionStatus");


const $ribbonTabs = $("#ribbonTabs");


const $ribbonPanels = $("#ribbonPanels");


const $ribbonDisplayButton = $("#ribbonDisplayButton");


const $ribbonDisplayMenu = $("#ribbonDisplayMenu");


const $fileSideMenu = $("#fileSideMenu");


const $statusText = $("#statusText");


const $workbookName = $("#workbookName");


const $workbookNameDisplay = $("#workbookNameDisplay");


const $workbookNameDisplayText = $("#workbookNameDisplayText");


const $workbookWindowBar = $("#workbookWindowBar");


const $warnings = $("#warnings");


const $appAlert = $("#appAlert");


const $selectionKeyboardSink = $("#selectionKeyboardSink");


const $formulaBar = $(".formula-bar");


const $formulaInput = $("#formulaInput");


const $formulaExpandButton = $("#formulaExpandButton");


const $formulaResizeHandle = $("#formulaResizeHandle");


const $selectedCellLabel = $("#selectedCellLabel");



let workbookNameMutationObserver = null;


let workbookNameResizeObserver = null;


let workbookNameDisplayRaf = 0;


let workbookNameMeasureContext = null;


const $exportHtmlButton = $("#exportHtmlButton");


const $toggleFormulasButton = $("#toggleFormulasButton");


const CanvasRenderer = {
  renderLayer: renderCanvasLayer,
  paint: paintSheetCanvas,
};



const SPECIAL_SHEET_INSERT_META = {
  macro: {
    type: "macro",
    label: "MS Excel 4.0 マクロ",
    nameBase: "Macro",
    sheetKind: "macro",
    codeObjectName: "Macro",
    statusNoun: "MS Excel 4.0 マクロ シート",
  },
  intlMacro: {
    type: "intlMacro",
    label: "国際対応マクロ シート",
    nameBase: "IntlMacro",
    sheetKind: "intlMacro",
    codeObjectName: "InternationalMacro",
    statusNoun: "国際対応マクロ シート",
  },
  dialog: {
    type: "dialog",
    label: "ダイアログ シート",
    nameBase: "Dialog",
    sheetKind: "dialog",
    codeObjectName: "DialogSheet",
    statusNoun: "ダイアログ シート",
  },
};



const FORMULA_REFERENCE_COLORS = [
  { border: "#2563eb", fill: "rgba(37, 99, 235, 0.08)" },
  { border: "#dc2626", fill: "rgba(220, 38, 38, 0.08)" },
  { border: "#16a34a", fill: "rgba(22, 163, 74, 0.08)" },
  { border: "#9333ea", fill: "rgba(147, 51, 234, 0.08)" },
  { border: "#ea580c", fill: "rgba(234, 88, 12, 0.08)" },
  { border: "#0891b2", fill: "rgba(8, 145, 178, 0.08)" },
];



const AUTO_FILL_MODES = new Set(["copy", "series", "formatting", "no-formatting", "date-day", "date-weekday", "date-month", "date-year", "linear-trend", "growth-trend"]);


const AUTO_FILL_DATE_MODES = new Set(["date-day", "date-weekday", "date-month", "date-year"]);


const AUTO_FILL_TREND_MODES = new Set(["linear-trend", "growth-trend"]);


const CWS_AGENT_MAX_RANGE_CELLS = 400;


const CWS_AGENT_MAX_FORMAT_CELLS = 50000;


const CWS_AGENT_MAX_COPY_CELLS = 2000;


const CWS_AGENT_MAX_SORT_CELLS = 2000;


const CWS_AGENT_MAX_FILTER_CELLS = 10000;


const CWS_AGENT_MAX_ADVANCED_FILTER_CELLS = 50000;


const CWS_AGENT_MAX_VALIDATION_CELLS = 2000;


const CWS_AGENT_MAX_DEDUP_CELLS = 10000;


const CWS_AGENT_MAX_FIND_REPLACE_CELLS = 10000;


const CWS_AGENT_MAX_TEXT_TO_COLUMNS_CELLS = 10000;


const CWS_AGENT_MAX_TABLE_CELLS = 50000;


const CWS_AGENT_MAX_CHART_CELLS = 10000;


const CWS_AGENT_MAX_PIVOT_CELLS = 50000;


const CWS_AGENT_MAX_CONSOLIDATE_CELLS = 50000;


const CWS_AGENT_MAX_FORECAST_CELLS = 10000;


const CWS_AGENT_MAX_DATA_TABLE_CELLS = 20000;


const CWS_AGENT_MAX_QUICK_ANALYSIS_CELLS = 10000;


const CWS_AGENT_MAX_FILTER_VALUES = 200;


const CWS_AGENT_MAX_FILTER_VALUE_TEXT = 500;


const CWS_AGENT_MAX_VALIDATION_OPTIONS = 200;


const CWS_AGENT_MAX_VALIDATION_OPTION_TEXT = 500;


const CWS_AGENT_MAX_HYPERLINK_TEXT = 2000;


const CWS_AGENT_MAX_NOTE_TEXT = 30000;


const CWS_AGENT_MAX_COMMENT_TEXT = 30000;


const CWS_AGENT_MAX_SPARKLINE_SOURCE_CELLS = 2000;


const CWS_AGENT_MAX_STRUCTURE_ROWS = 100;


const CWS_AGENT_MAX_STRUCTURE_COLUMNS = 100;


const CWS_AGENT_MIN_COLUMN_WIDTH = 8;


const CWS_AGENT_MAX_COLUMN_WIDTH = 600;


const CWS_AGENT_MIN_ROW_HEIGHT = 4;


const CWS_AGENT_MAX_ROW_HEIGHT = 300;


const CWS_AGENT_MAX_FREEZE_ROWS = 1000;


const CWS_AGENT_MAX_FREEZE_COLUMNS = 100;


const CWS_AGENT_MIN_CHART_SIZE = 160;


const CWS_AGENT_MAX_CHART_SIZE = 1200;


// Agent geometry is explicit; the mouse-drawing minimum must not enlarge icons.
const CWS_AGENT_MIN_SHAPE_SIZE = 1;


const CWS_AGENT_MAX_SHAPE_SIZE = 1600;


const CWS_AGENT_MAX_SHAPE_TEXT = 2000;


const CWS_AGENT_MAX_GOAL_SEEK_ABS_TARGET = 1e15;


const CWS_AGENT_MAX_WORKBOOK_TEMPLATE_SHEETS = 30;


const CWS_AGENT_MAX_WORKBOOK_TEMPLATE_CELLS = 5000;


const CWS_AGENT_MAX_OPS = 1000;


const CWS_AGENT_PROGRESSIVE_PREVIEW_DELAY_MS = 90;


const CWS_AGENT_BORDER_PRESETS = new Set(["none", "outside", "thick-outside", "all", "inside", "top", "right", "bottom", "left", "thick-bottom", "remove-outside"]);


const CWS_AGENT_CONTRACT_RANGE_OPS = new Set([
  "setRange",
  "setStyle",
  "setCellProtection",
  "clearCellProtection",
  "setBorders",
  "setNumberFormat",
  "sortRange",
  "setAutoFilter",
  "clearAutoFilter",
  "setFilterCriteria",
  "clearFilterCriteria",
  "advancedFilter",
  "removeDuplicates",
  "findReplace",
  "textToColumns",
  "createTable",
  "createChart",
  "createPivotTable",
  "createForecastSheet",
  "dataTable",
  "setQuickAnalysis",
  "clearQuickAnalysis",
  "setQuickAnalysisTotal",
  "setSubtotal",
  "clearSubtotals",
  "consolidateRanges",
  "setDataValidation",
  "clearDataValidation",
  "setHyperlink",
  "clearHyperlink",
  "setNote",
  "clearNote",
  "setComment",
  "appendComment",
  "clearComment",
  "setNamedRange",
  "clearRange",
  "mergeCells",
  "unmergeCells",
  "selectRange",
]);



const cwsAgentOperationSources = new WeakMap();



const LEGACY_XLS_BORDER_STYLE_MAP = {
  1: "thin",
  2: "medium",
  3: "dashed",
  4: "dotted",
  5: "thick",
  6: "double",
  7: "hair",
  8: "mediumDashed",
  9: "dashDot",
  10: "mediumDashDot",
  11: "dashDotDot",
  12: "mediumDashDotDot",
  13: "slantDashDot",
};



const LEGACY_XLS_BUILTIN_NUM_FMTS = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  5: '"$"#,##0_);("$"#,##0)',
  6: '"$"#,##0_);[Red]("$"#,##0)',
  7: '"$"#,##0.00_);("$"#,##0.00)',
  8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "m/d/yy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yy h:mm",
  37: "#,##0 ;(#,##0)",
  38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mmss.0",
  48: "##0.0E+0",
  49: "@",
};



const LEGACY_XLS_FILL_PATTERN_NONE = new Set([0]);



const LEGACY_XLS_CELL_RECORDS_WITH_XF = new Set([
  0x0002,
  0x0003,
  0x0004,
  0x0005,
  0x0006,
  0x00d6,
  0x00fd,
  0x0201,
  0x0203,
  0x0204,
  0x0205,
  0x0206,
  0x027e,
  0x0406,
]);



const CJK_FONT_FALLBACKS = [
  "Hiragino Sans",
  "Hiragino Kaku Gothic ProN",
  "Yu Gothic",
  "Yu Gothic UI",
  "Meiryo",
  "Noto Sans JP",
  "Noto Sans CJK JP",
  "Source Han Sans JP",
  "IPAexGothic",
  "TakaoPGothic",
  "sans-serif",
];



const EXCEL_FONT_FALLBACKS = {
  "ms pゴシック": ["MS Pゴシック", "MS PGothic", "ＭＳ Ｐゴシック", "MS P Gothic", "Yu Gothic", "Meiryo"],
  "ｍｓ ｐゴシック": ["ＭＳ Ｐゴシック", "MS PGothic", "MS P Gothic", "Yu Gothic", "Meiryo"],
  "ｍｓ pゴシック": ["ＭＳ Ｐゴシック", "MS Pゴシック", "MS PGothic", "MS P Gothic", "Yu Gothic", "Meiryo"],
  "ms pgothic": ["MS PGothic", "ＭＳ Ｐゴシック", "MS P Gothic", "Yu Gothic", "Meiryo"],
  "ms p gothic": ["MS P Gothic", "MS PGothic", "ＭＳ Ｐゴシック", "Yu Gothic", "Meiryo"],
  "ms gothic": ["MS Gothic", "ＭＳ ゴシック", "Yu Gothic", "Meiryo"],
  "ｍｓ ゴシック": ["ＭＳ ゴシック", "MS Gothic", "Yu Gothic", "Meiryo"],
  "ms ui gothic": ["MS UI Gothic", "MS PGothic", "Yu Gothic UI", "Meiryo"],
  "ms p明朝": ["MS P明朝", "MS PMincho", "ＭＳ Ｐ明朝", "Yu Mincho", "Hiragino Mincho ProN"],
  "ms pmincho": ["MS PMincho", "ＭＳ Ｐ明朝", "Yu Mincho", "Hiragino Mincho ProN"],
  "ヒラギノ角ゴ pron w3": ["ヒラギノ角ゴ ProN W3", "Hiragino Kaku Gothic ProN", "Hiragino Sans"],
  "hiragino sans": ["Hiragino Sans", "Hiragino Kaku Gothic ProN", "Noto Sans JP"],
  "hiragino kaku gothic pron": ["Hiragino Kaku Gothic ProN", "Hiragino Sans", "ヒラギノ角ゴ ProN W3"],
  "hiragino mincho pron": ["Hiragino Mincho ProN", "Yu Mincho", "Noto Serif CJK JP"],
  "noto sans cjk jp": ["Noto Sans CJK JP", "Noto Sans JP", "Hiragino Sans", "Yu Gothic"],
  "noto sans jp": ["Noto Sans JP", "Noto Sans CJK JP", "Hiragino Sans", "Yu Gothic"],
  "noto serif cjk jp": ["Noto Serif CJK JP", "Yu Mincho", "Hiragino Mincho ProN"],
  "source han sans jp": ["Source Han Sans JP", "Noto Sans CJK JP", "Noto Sans JP"],
  "yu gothic": ["Yu Gothic", "Yu Gothic UI", "Meiryo", "Hiragino Sans"],
  "yu gothic ui": ["Yu Gothic UI", "Yu Gothic", "Meiryo", "Hiragino Sans"],
  meiryo: ["Meiryo", "Yu Gothic", "Noto Sans JP", "Hiragino Sans"],
  calibri: ["Calibri", "Arial", "Helvetica Neue"],
};



const INDEXED_COLORS = {
  0: "#000000",
  1: "#FFFFFF",
  2: "#FF0000",
  3: "#00FF00",
  4: "#0000FF",
  5: "#FFFF00",
  6: "#FF00FF",
  7: "#00FFFF",
  8: "#000000",
  9: "#FFFFFF",
  10: "#FF0000",
  11: "#00FF00",
  12: "#0000FF",
  13: "#FFFF00",
  14: "#FF00FF",
  15: "#00FFFF",
  16: "#800000",
  17: "#008000",
  18: "#000080",
  19: "#808000",
  20: "#800080",
  21: "#008080",
  22: "#C0C0C0",
  23: "#808080",
  24: "#9999FF",
  25: "#993366",
  26: "#FFFFCC",
  27: "#CCFFFF",
  28: "#660066",
  29: "#FF8080",
  30: "#0066CC",
  31: "#CCCCFF",
  32: "#000080",
  33: "#FF00FF",
  34: "#FFFF00",
  35: "#00FFFF",
  36: "#800080",
  37: "#800000",
  38: "#008080",
  39: "#0000FF",
  40: "#00CCFF",
  41: "#CCFFFF",
  42: "#CCFFCC",
  43: "#FFFF99",
  44: "#99CCFF",
  45: "#FF99CC",
  46: "#CC99FF",
  47: "#FFCC99",
  48: "#3366FF",
  49: "#33CCCC",
  50: "#99CC00",
  51: "#FFCC00",
  52: "#FF9900",
  53: "#FF6600",
  54: "#666699",
  55: "#969696",
  56: "#003366",
  57: "#339966",
  58: "#003300",
  59: "#333300",
  60: "#993300",
  61: "#993366",
  62: "#333399",
  63: "#333333",
  64: "#000000",
};



const IMPORTED_RANGE_DIMENSION_EXPANSION_CELL_LIMIT = 20000;



const WEBSHEET_FORMULA_ENGINE_CACHEABLE_SPILL_NAMES = new Set([
  "UNIQUE",
  "SORT",
  "SORTBY",
  "FILTER",
  "SEQUENCE",
  "INDIRECT",
  "OFFSET",
  "INDEX",
  "TAKE",
  "DROP",
  "VSTACK",
  "HSTACK",
  "CHOOSEROWS",
  "CHOOSECOLS",
  "TOCOL",
  "TOROW",
  "WRAPROWS",
  "WRAPCOLS",
  "EXPAND",
  "TRANSPOSE",
  "CHOOSE",
  "TEXTSPLIT",
  "MAKEARRAY",
  "BYROW",
  "BYCOL",
  "MAP",
  "REDUCE",
  "SCAN",
]);



const FORMULA_ENGINE_STATIC_CELL_INFO_TYPES = new Set(["address", "col", "row"]);



const SHEET_INSERT_TEMPLATES = [
  { type: "worksheet", label: "ワークシート", icon: "▦" },
  { type: "chart", label: "グラフ", icon: "▧" },
  { type: "macro", label: "MS Excel 4.0 マクロ", icon: "▤" },
  { type: "intlMacro", label: "国際対応マクロ シート", icon: "▥" },
  { type: "dialog", label: "ダイアログ シート", icon: "▣" },
];



const TRANSLATION_CELL_LIMIT = 500;


const TRANSLATION_SCAN_LIMIT = 4000;


const TRANSLATION_PREVIEW_LIMIT = 80;


const TRANSLATION_LANGUAGES = [
  { value: "auto", label: "自動検出" },
  { value: "ja", label: "日本語" },
  { value: "zh-CN", label: "中国語（簡体字）" },
  { value: "en", label: "英語" },
  { value: "ko", label: "韓国語" },
];



const OFFLINE_TRANSLATION_DICTIONARIES = {
  "ja>zh-CN": {
    "スキルシート": "技能表",
    "氏名": "姓名",
    "年齢": "年龄",
    "性別": "性别",
    "経験年数": "经验年数",
    "最終学歴": "最高学历",
    "最寄駅": "最近车站",
    "資格・講座名": "资格・课程名",
    "資格": "资格",
    "講座名": "课程名",
    "取得日": "取得日",
    "日本語能力試験": "日语能力考试",
    "日本語": "日语",
    "英語": "英语",
    "中国語": "中文",
    "男性": "男性",
    "女性": "女性",
    "男": "男",
    "女": "女",
    "歳": "岁",
    "駅": "站",
    "線": "线",
    "評価": "评价",
    "指導可能": "可指导",
    "一人称で可能": "可独立完成",
    "サポートありで可能": "有支持可完成",
    "知見あり": "有相关知识",
    "実務経験なし": "无实务经验",
    "目的": "目的",
    "特徴": "特点",
    "生成AI": "生成式AI",
    "ローカルLLM": "本地LLM",
    "ワークフロー": "工作流",
    "ナレッジ": "知识",
    "知識": "知识",
    "連携": "联动",
    "実行": "执行",
    "管理": "管理",
    "基盤": "平台",
    "設計": "设计",
    "開発": "开发",
    "検証": "验证",
    "実装": "实现",
    "業務": "业务",
    "技術": "技术",
    "品質": "质量",
    "コスト": "成本",
    "納期": "交期",
    "課題": "课题",
    "判断": "判断",
    "経験": "经验",
    "プロジェクト": "项目",
    "チーム": "团队",
    "レビュー": "评审",
    "成果": "成果",
    "独立": "独立",
    "概要": "概要",
    "詳細": "详细",
    "備考": "备注",
  },
  "ja>en": {
    "スキルシート": "Skill Sheet",
    "氏名": "Name",
    "年齢": "Age",
    "性別": "Gender",
    "経験年数": "Years of Experience",
    "最終学歴": "Education",
    "最寄駅": "Nearest Station",
    "資格・講座名": "Certification / Course",
    "資格": "Certification",
    "取得日": "Date Acquired",
    "日本語能力試験": "Japanese Language Proficiency Test",
    "日本語": "Japanese",
    "英語": "English",
    "中国語": "Chinese",
    "男性": "Male",
    "女性": "Female",
    "男": "Male",
    "女": "Female",
    "歳": "years old",
    "駅": "Station",
    "線": "Line",
    "評価": "Rating",
    "目的": "Purpose",
    "特徴": "Features",
    "生成AI": "Generative AI",
    "ローカルLLM": "Local LLM",
    "ワークフロー": "Workflow",
    "知識": "Knowledge",
    "連携": "Integration",
    "実行": "Execution",
    "管理": "Management",
    "基盤": "Platform",
    "設計": "Design",
    "開発": "Development",
    "検証": "Verification",
    "実装": "Implementation",
    "業務": "Business",
    "技術": "Technology",
    "品質": "Quality",
    "コスト": "Cost",
    "納期": "Delivery",
    "課題": "Issue",
    "経験": "Experience",
    "プロジェクト": "Project",
    "チーム": "Team",
    "概要": "Summary",
    "詳細": "Details",
    "備考": "Notes",
  },
  "zh-CN>ja": {
    "技能表": "スキルシート",
    "姓名": "氏名",
    "年龄": "年齢",
    "性别": "性別",
    "经验年数": "経験年数",
    "最高学历": "最終学歴",
    "最近车站": "最寄駅",
    "资格": "資格",
    "课程名": "講座名",
    "获得日期": "取得日",
    "日语能力考试": "日本語能力試験",
    "日语": "日本語",
    "英语": "英語",
    "中文": "中国語",
    "岁": "歳",
    "站": "駅",
    "线": "線",
    "评价": "評価",
    "目的": "目的",
    "特点": "特徴",
    "生成式AI": "生成AI",
    "本地LLM": "ローカルLLM",
    "工作流": "ワークフロー",
    "知识": "知識",
    "联动": "連携",
    "执行": "実行",
    "管理": "管理",
    "平台": "基盤",
    "设计": "設計",
    "开发": "開発",
    "验证": "検証",
    "实现": "実装",
    "业务": "業務",
    "技术": "技術",
    "质量": "品質",
    "成本": "コスト",
    "交期": "納期",
    "课题": "課題",
    "项目": "プロジェクト",
    "团队": "チーム",
    "备注": "備考",
  },
  "en>ja": {
    "Skill Sheet": "スキルシート",
    "Name": "氏名",
    "Age": "年齢",
    "Gender": "性別",
    "Years of Experience": "経験年数",
    "Education": "最終学歴",
    "Nearest Station": "最寄駅",
    "Certification": "資格",
    "Date Acquired": "取得日",
    "Japanese": "日本語",
    "English": "英語",
    "Chinese": "中国語",
    "Male": "男",
    "Female": "女",
    "Rating": "評価",
    "Purpose": "目的",
    "Features": "特徴",
    "Generative AI": "生成AI",
    "Local LLM": "ローカルLLM",
    "Workflow": "ワークフロー",
    "Knowledge": "知識",
    "Integration": "連携",
    "Execution": "実行",
    "Management": "管理",
    "Platform": "基盤",
    "Design": "設計",
    "Development": "開発",
    "Verification": "検証",
    "Implementation": "実装",
    "Business": "業務",
    "Technology": "技術",
    "Quality": "品質",
    "Cost": "コスト",
    "Delivery": "納期",
    "Project": "プロジェクト",
    "Team": "チーム",
    "Summary": "概要",
    "Details": "詳細",
    "Notes": "備考",
  },
};



const QUICK_ANALYSIS_TABS = [
  { id: "format", label: "書式", icon: "paintbrush" },
  { id: "charts", label: "グラフ", icon: "bar-chart-3" },
  { id: "totals", label: "集計", icon: "sigma" },
  { id: "tables", label: "テーブル", icon: "table-2" },
  { id: "sparklines", label: "スパークライン", icon: "chart-line" },
];



const QUICK_ANALYSIS_ACTIONS = {
  format: [
    { action: "format-data-bars", label: "データ バー", icon: "bar-chart-3" },
    { action: "format-color-scale", label: "カラー スケール", icon: "palette" },
    { action: "format-icon-set", label: "アイコン セット", icon: "shapes" },
    { action: "format-greater-than", label: "指定の値より大きい", icon: "badge-plus" },
    { action: "format-duplicates", label: "重複する値", icon: "copy" },
    { action: "format-clear", label: "書式のクリア", icon: "eraser" },
  ],
  charts: [
    { action: "chart-column", label: "集合縦棒", icon: "bar-chart-3" },
    { action: "chart-bar", label: "横棒", icon: "chart-no-axes-column" },
    { action: "chart-line", label: "折れ線", icon: "chart-line" },
    { action: "chart-pie", label: "円", icon: "chart-pie" },
    { action: "chart-scatter", label: "散布図", icon: "scatter-chart" },
    { action: "chart-recommended", label: "その他のグラフ", icon: "sparkles" },
  ],
  totals: [
    { action: "total-sum", label: "合計", icon: "sigma" },
    { action: "total-average", label: "平均", icon: "sigma" },
    { action: "total-count", label: "データの個数", icon: "list-checks" },
    { action: "total-percent", label: "% 合計", icon: "hash" },
    { action: "total-running", label: "累計", icon: "arrow-up-down" },
  ],
  tables: [
    { action: "table", label: "テーブル", icon: "table-2" },
    { action: "pivot-table", label: "ピボットテーブル", icon: "table-properties" },
  ],
  sparklines: [
    { action: "sparkline-line", label: "折れ線", icon: "chart-line" },
    { action: "sparkline-column", label: "縦棒", icon: "bar-chart-3" },
    { action: "sparkline-winloss", label: "勝敗", icon: "arrow-up-down" },
  ],
};



const FORMAT_DIALOG_DATE_FORMATS = [
  "yyyy/m/d",
  "yyyy/mm/dd",
  "yy/m/d",
  "m/d",
  "m/d/yy",
  "d-mmm",
  "mmm-yy",
  "mmmm d, yyyy",
  "yyyy年m月d日",
  "yyyy\"年\"m\"月\"d\"日",
  "m月d日",
  "yyyy/m/d h:mm",
  "yyyy/mm/dd h:mm",
  "yyyy/m/d h:mm:ss",
  "h:mm",
  "h:mm:ss",
  "h:mm AM/PM",
];



const RIBBON_PANEL_FOCUSABLE_SELECTOR = [
  "button",
  "select",
  "input:not([type='hidden'])",
  "textarea",
  "[tabindex]:not([tabindex='-1'])",
].join(",");



const WORKBOOK_OPTIONAL_RIBBON_COMMANDS = new Set(["open", "open-web-version", "new-workbook", "help", "version"]);



let formulaEngineAutocompleteFunctionCache = null;



function formulaEngineAutocompleteFunctionDefinitions() {
  if (formulaEngineAutocompleteFunctionCache) return formulaEngineAutocompleteFunctionCache;
  let temporaryEngine = null;
  try {
    const hf = state.hf || HyperFormula.buildEmpty(hyperFormulaBuildConfig());
    if (!state.hf) temporaryEngine = hf;
    const names = typeof hf.getRegisteredFunctionNames === "function"
      ? hf.getRegisteredFunctionNames()
      : [];
    formulaEngineAutocompleteFunctionCache = [...new Set(names.map((name) => String(name || "").toUpperCase()))]
      .filter((name) => formulaAutocompleteShouldExposeEngineFunction(name))
      .sort((left, right) => left.localeCompare(right, "en"))
      .map((name, index) => {
        const metadata = formulaEngineFunctionMetadata(hf, name);
        return {
          kind: "function",
          name,
          label: name,
          category: formulaEngineFunctionCategory(name),
          syntax: formulaEngineFunctionSyntax(name, metadata),
          description: formulaEngineFunctionDescription(name),
          priority: 120 + index,
          template: (ref) => formulaDefaultTemplate(name, ref),
          engineSupported: true,
        };
      });
  } catch (error) {
    console.warn("Formula autocomplete engine function scan failed:", error);
    formulaEngineAutocompleteFunctionCache = [];
  } finally {
    try {
      temporaryEngine?.destroy?.();
    } catch (error) {
      console.warn("Formula autocomplete temporary engine cleanup failed:", error);
    }
  }
  return formulaEngineAutocompleteFunctionCache;
}



const FORMULA_COMPATIBILITY_FUNCTIONS = new Set([
  "BETADIST",
  "BETAINV",
  "BINOMDIST",
  "CHIDIST",
  "CHIDISTRT",
  "CHIINV",
  "CHIINVRT",
  "CHITEST",
  "CONFIDENCE",
  "COVAR",
  "COVARIANCEP",
  "COVARIANCES",
  "CRITBINOM",
  "EXPONDIST",
  "FDIST",
  "FDISTRT",
  "FINV",
  "FINVRT",
  "FORECAST",
  "FTEST",
  "GAMMADIST",
  "GAMMAINV",
  "HYPGEOMDIST",
  "LOGINV",
  "LOGNORMDIST",
  "LOGNORMINV",
  "MODE",
  "NEGBINOMDIST",
  "NORMDIST",
  "NORMINV",
  "NORMSDIST",
  "NORMSINV",
  "POISSON",
  "POISSONDIST",
  "PERCENTILE",
  "QUARTILE",
  "RANK",
  "SKEWP",
  "STDEV",
  "STDEVP",
  "STDEVS",
  "TINV",
  "TINV2T",
  "TDIST",
  "TDIST2T",
  "TDISTRT",
  "TTEST",
  "VAR",
  "VARP",
  "VARS",
  "WEIBULL",
  "WEIBULLDIST",
  "ZTEST",
]);



const FORMULA_ENGINE_DESCRIPTION_OVERRIDES = {
  AVERAGEA: "文字列や論理値も含めて平均を返します。",
  COUNTBLANK: "範囲内の空白セルの個数を返します。",
  MAXA: "文字列や論理値も含めて最大値を返します。",
  MINA: "文字列や論理値も含めて最小値を返します。",
  MAXIFS: "複数条件に一致する最大値を返します。",
  MINIFS: "複数条件に一致する最小値を返します。",
  "QUARTILE.INC": "データの四分位数を返します。",
  "QUARTILE.EXC": "0 と 1 を除いた四分位数を返します。",
  "PERCENTILE.EXC": "0 と 1 を除いた百分位数を返します。",
  PERCENTILE: "百分位数を返す互換性関数です。",
  "PERCENTRANK.INC": "データ内の値の百分率の順位を返します。",
  "PERCENTRANK.EXC": "0 と 1 を除いた百分率の順位を返します。",
  KURT: "データセットの尖度を返します。",
  MODE: "最頻値を返す互換性関数です。",
  "MODE.MULT": "複数の最頻値を縦方向の配列として返します。",
  RANK: "数値の順位を返す互換性関数です。",
  "RANK.AVG": "同順位の場合は平均順位を返します。",
  FREQUENCY: "値の頻度分布を縦方向の配列として返します。",
  "NORM.DIST": "正規分布の値を返します。",
  "NORM.INV": "正規分布の逆関数の値を返します。",
  "T.TEST": "t 検定の確率を返します。",
  "CHISQ.TEST": "カイ二乗検定の値を返します。",
  "F.TEST": "F 検定の結果を返します。",
  "Z.TEST": "z 検定の片側確率を返します。",
  GAMMA: "ガンマ関数の値を返します。",
  "GAMMA.DIST": "ガンマ分布の値を返します。",
  "GAMMA.INV": "ガンマ分布の累積分布関数の逆関数を返します。",
  GAMMALN: "ガンマ関数の自然対数を返します。",
  "GAMMALN.PRECISE": "ガンマ関数の自然対数を返します。",
  "BETA.DIST": "ベータ分布の値を返します。",
  "BETA.INV": "ベータ分布の累積分布関数の逆関数を返します。",
  "CHISQ.DIST": "カイ二乗分布の値を返します。",
  "CHISQ.INV": "カイ二乗分布の累積分布関数の逆関数を返します。",
  "WEIBULL.DIST": "ワイブル分布の値を返します。",
  "HYPGEOM.DIST": "超幾何分布の値を返します。",
  "CONFIDENCE.NORM": "正規分布を使用して母平均の信頼区間を返します。",
  "CONFIDENCE.T": "t 分布を使用して母平均の信頼区間を返します。",
  "F.DIST": "F 分布の値を返します。",
  "F.DIST.RT": "F 分布の右側確率を返します。",
  "F.INV": "F 分布の累積分布関数の逆関数を返します。",
  "F.INV.RT": "F 分布の右側確率の逆関数を返します。",
  "T.DIST": "t 分布の値を返します。",
  "T.DIST.RT": "t 分布の右側確率を返します。",
  "T.DIST.2T": "t 分布の両側確率を返します。",
  "T.INV": "t 分布の累積分布関数の逆関数を返します。",
  "T.INV.2T": "t 分布の両側確率の逆関数を返します。",
  "LOGNORM.DIST": "対数正規分布の値を返します。",
  "LOGNORM.INV": "対数正規分布の累積分布関数の逆関数を返します。",
  "NEGBINOM.DIST": "負の二項分布の値を返します。",
  ERF: "誤差関数の値を返します。",
  ERFC: "相補誤差関数の値を返します。",
  PHI: "標準正規分布の密度関数の値を返します。",
  GAUSS: "標準正規分布の累積分布から 0.5 を引いた値を返します。",
  COMPLEX: "実数係数と虚数係数から複素数を作成します。",
  IMARGUMENT: "複素数の偏角を返します。",
  IMEXP: "複素数の指数関数を返します。",
  IMLN: "複素数の自然対数を返します。",
  IMLOG10: "複素数の常用対数を返します。",
  IMLOG2: "複素数の 2 を底とする対数を返します。",
  IMPOWER: "複素数のべき乗を返します。",
  IMSQRT: "複素数の平方根を返します。",
  IMSIN: "複素数のサインを返します。",
  IMCOS: "複素数のコサインを返します。",
  IMTAN: "複素数のタンジェントを返します。",
  IMSEC: "複素数のセカントを返します。",
  IMCSC: "複素数のコセカントを返します。",
  IMCOT: "複素数のコタンジェントを返します。",
  IMSINH: "複素数の双曲線サインを返します。",
  IMCOSH: "複素数の双曲線コサインを返します。",
  IMSECH: "複素数の双曲線セカントを返します。",
  IMCSCH: "複素数の双曲線コセカントを返します。",
  BESSELI: "修正ベッセル関数 In(x) を返します。",
  BESSELJ: "ベッセル関数 Jn(x) を返します。",
  BESSELK: "修正ベッセル関数 Kn(x) を返します。",
  BESSELY: "ベッセル関数 Yn(x) を返します。",
  "ERF.PRECISE": "誤差関数の値を返します。",
  BITAND: "2 つの数値のビット積を返します。",
  BITOR: "2 つの数値のビット和を返します。",
  BITXOR: "2 つの数値の排他的ビット和を返します。",
  BITLSHIFT: "数値を左にビットシフトします。",
  BITRSHIFT: "数値を右にビットシフトします。",
  DEC2BIN: "10 進数を 2 進数に変換します。",
  DEC2OCT: "10 進数を 8 進数に変換します。",
  DEC2HEX: "10 進数を 16 進数に変換します。",
  BIN2DEC: "2 進数を 10 進数に変換します。",
  BIN2OCT: "2 進数を 8 進数に変換します。",
  BIN2HEX: "2 進数を 16 進数に変換します。",
  OCT2BIN: "8 進数を 2 進数に変換します。",
  OCT2DEC: "8 進数を 10 進数に変換します。",
  OCT2HEX: "8 進数を 16 進数に変換します。",
  HEX2BIN: "16 進数を 2 進数に変換します。",
  HEX2OCT: "16 進数を 8 進数に変換します。",
  HEX2DEC: "16 進数を 10 進数に変換します。",
  DELTA: "2 つの数値が等しいかどうかを返します。",
  GESTEP: "数値がしきい値以上かどうかを返します。",
  FVSCHEDULE: "一連の利率を適用した将来価値を返します。",
  XNPV: "日付付きキャッシュフローの正味現在価値を返します。",
  XIRR: "日付付きキャッシュフローの内部利益率を返します。",
  DB: "固定定率法を使用して、指定した期の減価償却費を返します。",
  DDB: "倍額定率法または指定した率で、指定した期の減価償却費を返します。",
  VDB: "可変定率法を使用して、指定した期間の減価償却費を返します。",
  AMORLINC: "各会計期の減価償却費を返します。",
  AMORDEGRC: "減価償却係数を使用して、各会計期の減価償却費を返します。",
  SLN: "定額法による 1 期あたりの減価償却費を返します。",
  SYD: "級数法による指定した期の減価償却費を返します。",
  DOLLARDE: "分数表示の価格を小数表示に変換します。",
  DOLLARFR: "小数表示の価格を分数表示に変換します。",
  RRI: "投資の成長に対する等価利率を返します。",
  PDURATION: "投資が指定した価値に達するまでの期間を返します。",
  ISPMT: "投資期間内の指定した期の利息を返します。",
  TBILLEQ: "短期国庫証券の債券換算利回りを返します。",
  TBILLPRICE: "短期国庫証券の額面 100 あたりの価格を返します。",
  TBILLYIELD: "短期国庫証券の利回りを返します。",
  DISC: "割引債の割引率を返します。",
  PRICEDISC: "割引債の額面 100 あたりの価格を返します。",
  YIELDDISC: "割引債の利回りを返します。",
  INTRATE: "満期一括払い証券の利率を返します。",
  RECEIVED: "満期時に受け取る金額を返します。",
  ACCRINT: "定期的に利息が支払われる証券の未収利息を返します。",
  ACCRINTM: "満期日に利息が支払われる証券の未収利息を返します。",
  PRICEMAT: "満期日に利息が支払われる証券の額面 100 あたりの価格を返します。",
  YIELDMAT: "満期日に利息が支払われる証券の利回りを返します。",
  COUPDAYBS: "利払期間の開始日から受渡日までの日数を返します。",
  COUPDAYS: "受渡日を含む利払期間の日数を返します。",
  COUPDAYSNC: "受渡日から次の利払日までの日数を返します。",
  COUPNCD: "受渡日後の次の利払日を返します。",
  COUPPCD: "受渡日前の前回利払日を返します。",
  COUPNUM: "受渡日から満期日までの利払回数を返します。",
  PRICE: "定期的に利息が支払われる証券の額面 100 あたりの価格を返します。",
  YIELD: "定期的に利息が支払われる証券の利回りを返します。",
  DURATION: "定期的に利息が支払われる証券のマコーレー デュレーションを返します。",
  MDURATION: "修正マコーレー デュレーションを返します。",
  SQRTPI: "数値に円周率を掛けた値の平方根を返します。",
  LOOKUP: "1 行または 1 列の範囲から値を検索します。",
  ROW: "参照の行番号を返します。",
  COLUMN: "参照の列番号を返します。",
  ROWS: "参照または配列の行数を返します。",
  COLUMNS: "参照または配列の列数を返します。",
  MMULT: "2 つの配列の行列積を返します。",
  MDETERM: "配列の行列式を返します。",
  MINVERSE: "配列の逆行列を返します。",
  MUNIT: "指定した次元の単位行列を返します。",
  HYPERLINK: "リンクを作成します。",
  T: "値が文字列ならその文字列を返します。",
  DAYS: "2 つの日付の間の日数を返します。",
  DAYS360: "1 年を 360 日として日数を返します。",
  EDATE: "指定月数だけ前後した日付を返します。",
  DOLLAR: "数値を通貨形式の文字列に変換します。",
  FIXED: "数値を固定小数点形式の文字列に変換します。",
  NUMBERVALUE: "指定した区切り記号で文字列を数値に変換します。",
  SIGN: "数値の符号を返します。",
  PI: "円周率 π を返します。",
  SIN: "角度のサインを返します。",
  COS: "角度のコサインを返します。",
  TAN: "角度のタンジェントを返します。",
  ASIN: "アークサインを返します。",
  ACOS: "アークコサインを返します。",
  ATAN: "アークタンジェントを返します。",
  DEGREES: "ラジアンを度に変換します。",
  RADIANS: "度をラジアンに変換します。",
  ATAN2: "指定した x 座標と y 座標のアークタンジェントを返します。",
  SINH: "双曲線正弦を返します。",
  COSH: "双曲線余弦を返します。",
  TANH: "双曲線正接を返します。",
  ASINH: "双曲線逆正弦を返します。",
  ACOSH: "双曲線逆余弦を返します。",
  ATANH: "双曲線逆正接を返します。",
  COT: "コタンジェントを返します。",
  SEC: "セカントを返します。",
  CSC: "コセカントを返します。",
  ACOT: "アークコタンジェントを返します。",
  COTH: "双曲線コタンジェントを返します。",
  SECH: "双曲線セカントを返します。",
  CSCH: "双曲線コセカントを返します。",
  SERIESSUM: "べき級数の和を返します。",
  SUMSQ: "引数の平方和を返します。",
  SUMX2MY2: "2 つの配列で対応する値の平方差の合計を返します。",
  SUMX2PY2: "2 つの配列で対応する値の平方和の合計を返します。",
  SUMXMY2: "2 つの配列で対応する値の差の平方和を返します。",
  EVEN: "数値を最も近い偶数に切り上げます。",
  ODD: "数値を最も近い奇数に切り上げます。",
  "ISO.CEILING": "ISO 規格に従って基準値の倍数に切り上げます。",
  LN: "自然対数を返します。",
  LOG: "指定した底の対数を返します。",
  LOG10: "10 を底とする対数を返します。",
  EXP: "e のべき乗を返します。",
  FACT: "数値の階乗を返します。",
  FACTDOUBLE: "数値の二重階乗を返します。",
  COMBIN: "組み合わせの数を返します。",
  COMBINA: "重複を許す組み合わせの数を返します。",
  PERMUT: "順列の数を返します。",
  PERMUTATIONA: "重複を許す順列の数を返します。",
  MULTINOMIAL: "多項係数を返します。",
  GCD: "最大公約数を返します。",
  LCM: "最小公倍数を返します。",
  PV: "投資の現在価値を返します。",
  FV: "投資の将来価値を返します。",
  IPMT: "指定期間の利息支払額を返します。",
  PPMT: "指定期間の元金支払額を返します。",
  MIRR: "修正内部利益率を返します。",
  ISBLANK: "参照が空白かどうかを返します。",
  ISERROR: "値がエラーかどうかを返します。",
  ISNUMBER: "値が数値かどうかを返します。",
  ISTEXT: "値が文字列かどうかを返します。",
  ISNA: "値が #N/A かどうかを返します。",
  "ERROR.TYPE": "エラー値に対応する番号を返します。",
  FORMULATEXT: "参照セルの数式を文字列として返します。",
  ISFORMULA: "参照セルに数式が含まれるかどうかを返します。",
  ISREF: "値が参照かどうかを返します。",
  SHEET: "参照先または指定したシート名のシート番号を返します。",
  SHEETS: "参照またはブックに含まれるシート数を返します。",
  CELL: "セルの書式、位置、内容に関する情報を返します。",
  INFO: "現在の環境に関する情報を返します。",
  DSUM: "条件に一致するデータベース列の合計を返します。",
  DCOUNT: "条件に一致する数値セルの個数を返します。",
  DCOUNTA: "条件に一致する空白でないセルの個数を返します。",
  DAVERAGE: "条件に一致するデータベース列の平均を返します。",
  DMAX: "条件に一致するデータベース列の最大値を返します。",
  DMIN: "条件に一致するデータベース列の最小値を返します。",
  DPRODUCT: "条件に一致するデータベース列の積を返します。",
  DSTDEV: "条件に一致するデータベース列の標本標準偏差を返します。",
  DSTDEVP: "条件に一致するデータベース列の母集団標準偏差を返します。",
  DVAR: "条件に一致するデータベース列の標本分散を返します。",
  DVARP: "条件に一致するデータベース列の母集団分散を返します。",
  BETADIST: "Excel 2010 より前の形式でベータ分布の値を返します。",
  BETAINV: "Excel 2010 より前の形式でベータ分布の逆関数を返します。",
  BINOMDIST: "Excel 2010 より前の形式で二項分布の値を返します。",
  CHIDIST: "Excel 2010 より前の形式でカイ二乗分布の右側確率を返します。",
  CHIDISTRT: "Excel 2010 より前の形式でカイ二乗分布の右側確率を返します。",
  CHIINV: "Excel 2010 より前の形式でカイ二乗分布の右側確率の逆関数を返します。",
  CHIINVRT: "Excel 2010 より前の形式でカイ二乗分布の右側確率の逆関数を返します。",
  CHITEST: "Excel 2010 より前の形式で独立性の検定を返します。",
  CONFIDENCE: "Excel 2010 より前の形式で母平均の信頼区間を返します。",
  COVAR: "Excel 2010 より前の形式で共分散を返します。",
  COVARIANCEP: "母集団共分散を返す互換性関数です。",
  COVARIANCES: "標本共分散を返す互換性関数です。",
  CRITBINOM: "累積二項分布が基準値以上になる最小値を返します。",
  EXPONDIST: "Excel 2010 より前の形式で指数分布を返します。",
  FDIST: "Excel 2010 より前の形式で F 分布の右側確率を返します。",
  FDISTRT: "Excel 2010 より前の形式で F 分布の右側確率を返します。",
  FINV: "Excel 2010 より前の形式で F 分布の右側確率の逆関数を返します。",
  FINVRT: "Excel 2010 より前の形式で F 分布の右側確率の逆関数を返します。",
  FTEST: "Excel 2010 より前の形式で F 検定の確率を返します。",
  GAMMADIST: "Excel 2010 より前の形式でガンマ分布を返します。",
  GAMMAINV: "Excel 2010 より前の形式でガンマ分布の逆関数を返します。",
  HYPGEOMDIST: "Excel 2010 より前の形式で超幾何分布を返します。",
  LOGINV: "Excel 2010 より前の形式で対数正規分布の逆関数を返します。",
  LOGNORMDIST: "Excel 2010 より前の形式で対数正規分布を返します。",
  LOGNORMINV: "対数正規分布の逆関数を返す互換性関数です。",
  NEGBINOMDIST: "Excel 2010 より前の形式で負の二項分布を返します。",
  NORMDIST: "Excel 2010 より前の形式で正規分布を返します。",
  NORMINV: "Excel 2010 より前の形式で正規分布の逆関数を返します。",
  NORMSDIST: "Excel 2010 より前の形式で標準正規分布を返します。",
  NORMSINV: "Excel 2010 より前の形式で標準正規分布の逆関数を返します。",
  POISSON: "Excel 2010 より前の形式でポアソン分布を返します。",
  POISSONDIST: "ポアソン分布を返す互換性関数です。",
  STDEV: "Excel 2010 より前の形式で標本に基づく標準偏差を返します。",
  STDEVP: "Excel 2010 より前の形式で母集団全体に基づく標準偏差を返します。",
  TINV: "Excel 2010 より前の形式で t 分布の逆関数を返します。",
  TINV2T: "Excel 2010 より前の形式で t 分布の両側逆関数を返します。",
  TDIST: "Excel 2010 より前の形式で t 分布を返します。",
  TDIST2T: "Excel 2010 より前の形式で t 分布の両側確率を返します。",
  TDISTRT: "Excel 2010 より前の形式で t 分布の右側確率を返します。",
  TTEST: "Excel 2010 より前の形式で t 検定の確率を返します。",
  VAR: "Excel 2010 より前の形式で標本に基づく分散を返します。",
  VARP: "Excel 2010 より前の形式で母集団全体に基づく分散を返します。",
  WEIBULL: "Excel 2010 より前の形式でワイブル分布を返します。",
  WEIBULLDIST: "ワイブル分布を返す互換性関数です。",
  ZTEST: "Excel 2010 より前の形式で z 検定の片側確率を返します。",
  FORECAST: "Excel 2010 より前の形式で線形予測値を返します。",
  "ERFC.PRECISE": "相補誤差関数の値を返します。",
};



const FORMULA_ENGINE_SYNTAX_OVERRIDES = {
  COUNTBLANK: "COUNTBLANK(range)",
  AVERAGEA: "AVERAGEA(value1, [value2], ...)",
  MAXA: "MAXA(value1, [value2], ...)",
  MINA: "MINA(value1, [value2], ...)",
  MAXIFS: "MAXIFS(max_range, criteria_range1, criteria1, ...)",
  MINIFS: "MINIFS(min_range, criteria_range1, criteria1, ...)",
  "QUARTILE.INC": "QUARTILE.INC(array, quart)",
  "QUARTILE.EXC": "QUARTILE.EXC(array, quart)",
  QUARTILE: "QUARTILE(array, quart)",
  "PERCENTILE.EXC": "PERCENTILE.EXC(array, k)",
  PERCENTILE: "PERCENTILE(array, k)",
  "PERCENTRANK.INC": "PERCENTRANK.INC(array, x, [significance])",
  "PERCENTRANK.EXC": "PERCENTRANK.EXC(array, x, [significance])",
  PROB: "PROB(x_range, prob_range, lower_limit, [upper_limit])",
  KURT: "KURT(number1, [number2], ...)",
  MODE: "MODE(number1, [number2], ...)",
  "MODE.MULT": "MODE.MULT(number1, [number2], ...)",
  RANK: "RANK(number, ref, [order])",
  "RANK.AVG": "RANK.AVG(number, ref, [order])",
  "PERCENTILE.INC": "PERCENTILE.INC(array, k)",
  FREQUENCY: "FREQUENCY(data_array, bins_array)",
  TRIMMEAN: "TRIMMEAN(array, percent)",
  PEARSON: "PEARSON(array1, array2)",
  RSQ: "RSQ(known_y's, known_x's)",
  STEYX: "STEYX(known_y's, known_x's)",
  SKEW: "SKEW(number1, [number2], ...)",
  "SKEW.P": "SKEW.P(number1, [number2], ...)",
  "NORM.DIST": "NORM.DIST(x, mean, standard_dev, cumulative)",
  "NORM.INV": "NORM.INV(probability, mean, standard_dev)",
  "T.TEST": "T.TEST(array1, array2, tails, type)",
  "CHISQ.TEST": "CHISQ.TEST(actual_range, expected_range)",
  "F.TEST": "F.TEST(array1, array2)",
  "Z.TEST": "Z.TEST(array, x, [sigma])",
  GAMMA: "GAMMA(number)",
  "GAMMA.DIST": "GAMMA.DIST(x, alpha, beta, cumulative)",
  "GAMMA.INV": "GAMMA.INV(probability, alpha, beta)",
  GAMMALN: "GAMMALN(x)",
  "GAMMALN.PRECISE": "GAMMALN.PRECISE(x)",
  "BETA.DIST": "BETA.DIST(x, alpha, beta, cumulative, [A], [B])",
  "BETA.INV": "BETA.INV(probability, alpha, beta, [A], [B])",
  "BINOM.DIST.RANGE": "BINOM.DIST.RANGE(trials, probability_s, number_s, [number_s2])",
  "CHISQ.DIST": "CHISQ.DIST(x, deg_freedom, cumulative)",
  "CHISQ.INV": "CHISQ.INV(probability, deg_freedom)",
  "WEIBULL.DIST": "WEIBULL.DIST(x, alpha, beta, cumulative)",
  "HYPGEOM.DIST": "HYPGEOM.DIST(sample_s, number_sample, population_s, number_pop, cumulative)",
  "CONFIDENCE.NORM": "CONFIDENCE.NORM(alpha, standard_dev, size)",
  "CONFIDENCE.T": "CONFIDENCE.T(alpha, standard_dev, size)",
  "F.DIST": "F.DIST(x, deg_freedom1, deg_freedom2, cumulative)",
  "F.DIST.RT": "F.DIST.RT(x, deg_freedom1, deg_freedom2)",
  "F.INV": "F.INV(probability, deg_freedom1, deg_freedom2)",
  "F.INV.RT": "F.INV.RT(probability, deg_freedom1, deg_freedom2)",
  "T.DIST": "T.DIST(x, deg_freedom, cumulative)",
  "T.DIST.RT": "T.DIST.RT(x, deg_freedom)",
  "T.DIST.2T": "T.DIST.2T(x, deg_freedom)",
  "T.INV": "T.INV(probability, deg_freedom)",
  "T.INV.2T": "T.INV.2T(probability, deg_freedom)",
  "LOGNORM.DIST": "LOGNORM.DIST(x, mean, standard_dev, cumulative)",
  "LOGNORM.INV": "LOGNORM.INV(probability, mean, standard_dev)",
  "NEGBINOM.DIST": "NEGBINOM.DIST(number_f, number_s, probability_s, cumulative)",
  ERF: "ERF(lower_limit, [upper_limit])",
  ERFC: "ERFC(x)",
  PHI: "PHI(x)",
  GAUSS: "GAUSS(z)",
  BITAND: "BITAND(number1, number2)",
  BITOR: "BITOR(number1, number2)",
  BITXOR: "BITXOR(number1, number2)",
  BITLSHIFT: "BITLSHIFT(number, shift_amount)",
  BITRSHIFT: "BITRSHIFT(number, shift_amount)",
  DEC2BIN: "DEC2BIN(number, [places])",
  DEC2OCT: "DEC2OCT(number, [places])",
  DEC2HEX: "DEC2HEX(number, [places])",
  BIN2DEC: "BIN2DEC(number)",
  BIN2OCT: "BIN2OCT(number, [places])",
  BIN2HEX: "BIN2HEX(number, [places])",
  OCT2BIN: "OCT2BIN(number, [places])",
  OCT2DEC: "OCT2DEC(number)",
  OCT2HEX: "OCT2HEX(number, [places])",
  HEX2BIN: "HEX2BIN(number, [places])",
  HEX2OCT: "HEX2OCT(number, [places])",
  HEX2DEC: "HEX2DEC(number)",
  DELTA: "DELTA(number1, [number2])",
  GESTEP: "GESTEP(number, [step])",
  BETADIST: "BETADIST(x, alpha, beta, [A], [B])",
  BETAINV: "BETAINV(probability, alpha, beta, [A], [B])",
  BINOMDIST: "BINOMDIST(number_s, trials, probability_s, cumulative)",
  CHIDIST: "CHIDIST(x, deg_freedom)",
  CHIDISTRT: "CHIDISTRT(x, deg_freedom)",
  CHIINV: "CHIINV(probability, deg_freedom)",
  CHIINVRT: "CHIINVRT(probability, deg_freedom)",
  CHITEST: "CHITEST(actual_range, expected_range)",
  CONFIDENCE: "CONFIDENCE(alpha, standard_dev, size)",
  COVAR: "COVAR(array1, array2)",
  COVARIANCEP: "COVARIANCEP(array1, array2)",
  COVARIANCES: "COVARIANCES(array1, array2)",
  CRITBINOM: "CRITBINOM(trials, probability_s, alpha)",
  EXPONDIST: "EXPONDIST(x, lambda, cumulative)",
  FDIST: "FDIST(x, deg_freedom1, deg_freedom2)",
  FDISTRT: "FDISTRT(x, deg_freedom1, deg_freedom2)",
  FINV: "FINV(probability, deg_freedom1, deg_freedom2)",
  FINVRT: "FINVRT(probability, deg_freedom1, deg_freedom2)",
  FTEST: "FTEST(array1, array2)",
  GAMMADIST: "GAMMADIST(x, alpha, beta, cumulative)",
  GAMMAINV: "GAMMAINV(probability, alpha, beta)",
  HYPGEOMDIST: "HYPGEOMDIST(sample_s, number_sample, population_s, number_pop)",
  LOGINV: "LOGINV(probability, mean, standard_dev)",
  LOGNORMDIST: "LOGNORMDIST(x, mean, standard_dev)",
  LOGNORMINV: "LOGNORMINV(probability, mean, standard_dev)",
  NEGBINOMDIST: "NEGBINOMDIST(number_f, number_s, probability_s)",
  NORMDIST: "NORMDIST(x, mean, standard_dev, cumulative)",
  NORMINV: "NORMINV(probability, mean, standard_dev)",
  NORMSDIST: "NORMSDIST(z)",
  NORMSINV: "NORMSINV(probability)",
  POISSON: "POISSON(x, mean, cumulative)",
  POISSONDIST: "POISSONDIST(x, mean, cumulative)",
  STDEV: "STDEV(number1, [number2], ...)",
  STDEVP: "STDEVP(number1, [number2], ...)",
  TINV: "TINV(probability, deg_freedom)",
  TINV2T: "TINV2T(probability, deg_freedom)",
  TDIST: "TDIST(x, deg_freedom, tails)",
  TDIST2T: "TDIST2T(x, deg_freedom)",
  TDISTRT: "TDISTRT(x, deg_freedom)",
  TTEST: "TTEST(array1, array2, tails, type)",
  VAR: "VAR(number1, [number2], ...)",
  VARP: "VARP(number1, [number2], ...)",
  WEIBULL: "WEIBULL(x, alpha, beta, cumulative)",
  WEIBULLDIST: "WEIBULLDIST(x, alpha, beta, cumulative)",
  ZTEST: "ZTEST(array, x, [sigma])",
  SLOPE: "SLOPE(known_y's, known_x's)",
  LINEST: "LINEST(known_y's, [known_x's], [const], [stats])",
  LOGEST: "LOGEST(known_y's, [known_x's], [const], [stats])",
  INTERCEPT: "INTERCEPT(known_y's, known_x's)",
  "FORECAST.LINEAR": "FORECAST.LINEAR(x, known_y's, known_x's)",
  TREND: "TREND(known_y's, [known_x's], [new_x's], [const])",
  GROWTH: "GROWTH(known_y's, [known_x's], [new_x's], [const])",
  FORECAST: "FORECAST(x, known_y's, known_x's)",
  LOOKUP: "LOOKUP(lookup_value, lookup_vector, [result_vector])",
  AGGREGATE: "AGGREGATE(function_num, options, ref1, [ref2], ...)",
  SORTBY: "SORTBY(array, by_array1, [sort_order1], ...)",
  LET: "LET(name1, name_value1, calculation_or_name2, [name_value2], ...)",
  LAMBDA: "LAMBDA(parameter1, [parameter2], ..., calculation)",
  MAKEARRAY: "MAKEARRAY(rows, cols, lambda)",
  BYROW: "BYROW(array, lambda)",
  BYCOL: "BYCOL(array, lambda)",
  MAP: "MAP(array1, lambda_or_array2, ...)",
  REDUCE: "REDUCE([initial_value], array, lambda)",
  SCAN: "SCAN([initial_value], array, lambda)",
  SUMPRODUCT: "SUMPRODUCT(array1, [array2], ...)",
  NETWORKDAYS: "NETWORKDAYS(start_date, end_date, [holidays])",
  "NETWORKDAYS.INTL": "NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])",
  WORKDAY: "WORKDAY(start_date, days, [holidays])",
  "WORKDAY.INTL": "WORKDAY.INTL(start_date, days, [weekend], [holidays])",
  ISOWEEKNUM: "ISOWEEKNUM(date)",
  DAYS: "DAYS(end_date, start_date)",
  DAYS360: "DAYS360(start_date, end_date, [method])",
  EDATE: "EDATE(start_date, months)",
  DATEVALUE: "DATEVALUE(date_text)",
  TIMEVALUE: "TIMEVALUE(time_text)",
  YEARFRAC: "YEARFRAC(start_date, end_date, [basis])",
  ADDRESS: "ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])",
  ROW: "ROW([reference])",
  COLUMN: "COLUMN([reference])",
  ROWS: "ROWS(array)",
  COLUMNS: "COLUMNS(array)",
  SHEET: "SHEET([value])",
  SHEETS: "SHEETS([reference])",
  FORMULATEXT: "FORMULATEXT(reference)",
  ISFORMULA: "ISFORMULA(reference)",
  ISREF: "ISREF(value)",
  CELL: "CELL(info_type, [reference])",
  INFO: "INFO(type_text)",
  T: "T(value)",
  TRANSPOSE: "TRANSPOSE(array)",
  HYPERLINK: "HYPERLINK(link_location, [friendly_name])",
  MMULT: "MMULT(array1, array2)",
  MDETERM: "MDETERM(array)",
  MINVERSE: "MINVERSE(array)",
  MUNIT: "MUNIT(dimension)",
  CEILING: "CEILING(number, significance)",
  "CEILING.MATH": "CEILING.MATH(number, [significance], [mode])",
  "CEILING.PRECISE": "CEILING.PRECISE(number, [significance])",
  FLOOR: "FLOOR(number, significance)",
  "FLOOR.MATH": "FLOOR.MATH(number, [significance], [mode])",
  "FLOOR.PRECISE": "FLOOR.PRECISE(number, [significance])",
  MROUND: "MROUND(number, multiple)",
  POWER: "POWER(number, power)",
  SQRT: "SQRT(number)",
  SQRTPI: "SQRTPI(number)",
  TRUNC: "TRUNC(number, [num_digits])",
  QUOTIENT: "QUOTIENT(numerator, denominator)",
  SIGN: "SIGN(number)",
  PI: "PI()",
  SIN: "SIN(number)",
  COS: "COS(number)",
  TAN: "TAN(number)",
  ASIN: "ASIN(number)",
  ACOS: "ACOS(number)",
  ATAN: "ATAN(number)",
  DEGREES: "DEGREES(angle)",
  RADIANS: "RADIANS(angle)",
  ATAN2: "ATAN2(x_num, y_num)",
  SINH: "SINH(number)",
  COSH: "COSH(number)",
  TANH: "TANH(number)",
  ASINH: "ASINH(number)",
  ACOSH: "ACOSH(number)",
  ATANH: "ATANH(number)",
  COT: "COT(number)",
  SEC: "SEC(number)",
  CSC: "CSC(number)",
  ACOT: "ACOT(number)",
  COTH: "COTH(number)",
  SECH: "SECH(number)",
  CSCH: "CSCH(number)",
  SERIESSUM: "SERIESSUM(x, n, m, coefficients)",
  SUMSQ: "SUMSQ(number1, [number2], ...)",
  SUMX2MY2: "SUMX2MY2(array_x, array_y)",
  SUMX2PY2: "SUMX2PY2(array_x, array_y)",
  SUMXMY2: "SUMXMY2(array_x, array_y)",
  EVEN: "EVEN(number)",
  ODD: "ODD(number)",
  "ISO.CEILING": "ISO.CEILING(number, [significance])",
  LN: "LN(number)",
  LOG: "LOG(number, [base])",
  LOG10: "LOG10(number)",
  EXP: "EXP(number)",
  FACT: "FACT(number)",
  FACTDOUBLE: "FACTDOUBLE(number)",
  COMBIN: "COMBIN(number, number_chosen)",
  COMBINA: "COMBINA(number, number_chosen)",
  PERMUT: "PERMUT(number, number_chosen)",
  PERMUTATIONA: "PERMUTATIONA(number, number_chosen)",
  MULTINOMIAL: "MULTINOMIAL(number1, [number2], ...)",
  GCD: "GCD(number1, [number2], ...)",
  LCM: "LCM(number1, [number2], ...)",
  ROMAN: "ROMAN(number, [form])",
  ARABIC: "ARABIC(text)",
  BASE: "BASE(number, radix, [min_length])",
  DECIMAL: "DECIMAL(text, radix)",
  DOLLAR: "DOLLAR(number, [decimals])",
  FIXED: "FIXED(number, [decimals], [no_commas])",
  NUMBERVALUE: "NUMBERVALUE(text, [decimal_separator], [group_separator])",
  TEXTSPLIT: "TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with])",
  TEXTBEFORE: "TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])",
  TEXTAFTER: "TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])",
  REPLACE: "REPLACE(old_text, start_num, num_chars, new_text)",
  PROPER: "PROPER(text)",
  EXACT: "EXACT(text1, text2)",
  CHAR: "CHAR(number)",
  CODE: "CODE(text)",
  CLEAN: "CLEAN(text)",
  REPT: "REPT(text, number_times)",
  IPMT: "IPMT(rate, per, nper, pv, [fv], [type])",
  PPMT: "PPMT(rate, per, nper, pv, [fv], [type])",
  MIRR: "MIRR(values, finance_rate, reinvest_rate)",
  FVSCHEDULE: "FVSCHEDULE(principal, schedule)",
  XNPV: "XNPV(rate, values, dates)",
  XIRR: "XIRR(values, dates, [guess])",
  DB: "DB(cost, salvage, life, period, [month])",
  DDB: "DDB(cost, salvage, life, period, [factor])",
  VDB: "VDB(cost, salvage, life, start_period, end_period, [factor], [no_switch])",
  AMORLINC: "AMORLINC(cost, date_purchased, first_period, salvage, period, rate, [basis])",
  AMORDEGRC: "AMORDEGRC(cost, date_purchased, first_period, salvage, period, rate, [basis])",
  SLN: "SLN(cost, salvage, life)",
  SYD: "SYD(cost, salvage, life, per)",
  DOLLARDE: "DOLLARDE(fractional_dollar, fraction)",
  DOLLARFR: "DOLLARFR(decimal_dollar, fraction)",
  RRI: "RRI(nper, pv, fv)",
  PDURATION: "PDURATION(rate, pv, fv)",
  ISPMT: "ISPMT(rate, per, nper, pv)",
  TBILLEQ: "TBILLEQ(settlement, maturity, discount)",
  TBILLPRICE: "TBILLPRICE(settlement, maturity, discount)",
  TBILLYIELD: "TBILLYIELD(settlement, maturity, pr)",
  DISC: "DISC(settlement, maturity, pr, redemption, [basis])",
  PRICEDISC: "PRICEDISC(settlement, maturity, discount, redemption, [basis])",
  YIELDDISC: "YIELDDISC(settlement, maturity, pr, redemption, [basis])",
  INTRATE: "INTRATE(settlement, maturity, investment, redemption, [basis])",
  RECEIVED: "RECEIVED(settlement, maturity, investment, discount, [basis])",
  ACCRINT: "ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method])",
  ACCRINTM: "ACCRINTM(issue, settlement, rate, par, [basis])",
  PRICEMAT: "PRICEMAT(settlement, maturity, issue, rate, yld, [basis])",
  YIELDMAT: "YIELDMAT(settlement, maturity, issue, rate, pr, [basis])",
  COUPDAYBS: "COUPDAYBS(settlement, maturity, frequency, [basis])",
  COUPDAYS: "COUPDAYS(settlement, maturity, frequency, [basis])",
  COUPDAYSNC: "COUPDAYSNC(settlement, maturity, frequency, [basis])",
  COUPNCD: "COUPNCD(settlement, maturity, frequency, [basis])",
  COUPPCD: "COUPPCD(settlement, maturity, frequency, [basis])",
  COUPNUM: "COUPNUM(settlement, maturity, frequency, [basis])",
  PRICE: "PRICE(settlement, maturity, rate, yld, redemption, frequency, [basis])",
  YIELD: "YIELD(settlement, maturity, rate, pr, redemption, frequency, [basis])",
  DURATION: "DURATION(settlement, maturity, coupon, yld, frequency, [basis])",
  MDURATION: "MDURATION(settlement, maturity, coupon, yld, frequency, [basis])",
  ISBLANK: "ISBLANK(value)",
  ISERROR: "ISERROR(value)",
  ISNUMBER: "ISNUMBER(value)",
  ISTEXT: "ISTEXT(value)",
  ISNA: "ISNA(value)",
  N: "N(value)",
  "ERROR.TYPE": "ERROR.TYPE(error_val)",
  IMABS: "IMABS(inumber)",
  IMSUM: "IMSUM(inumber1, [inumber2], ...)",
  IMPRODUCT: "IMPRODUCT(inumber1, [inumber2], ...)",
  IMREAL: "IMREAL(inumber)",
  IMAGINARY: "IMAGINARY(inumber)",
  COMPLEX: "COMPLEX(real_num, i_num, [suffix])",
  IMARGUMENT: "IMARGUMENT(inumber)",
  IMEXP: "IMEXP(inumber)",
  IMLN: "IMLN(inumber)",
  IMLOG10: "IMLOG10(inumber)",
  IMLOG2: "IMLOG2(inumber)",
  IMPOWER: "IMPOWER(inumber, number)",
  IMSQRT: "IMSQRT(inumber)",
  IMSIN: "IMSIN(inumber)",
  IMCOS: "IMCOS(inumber)",
  IMTAN: "IMTAN(inumber)",
  IMSEC: "IMSEC(inumber)",
  IMCSC: "IMCSC(inumber)",
  IMCOT: "IMCOT(inumber)",
  IMSINH: "IMSINH(inumber)",
  IMCOSH: "IMCOSH(inumber)",
  IMSECH: "IMSECH(inumber)",
  IMCSCH: "IMCSCH(inumber)",
  BESSELI: "BESSELI(x, n)",
  BESSELJ: "BESSELJ(x, n)",
  BESSELK: "BESSELK(x, n)",
  BESSELY: "BESSELY(x, n)",
  CONVERT: "CONVERT(number, from_unit, to_unit)",
  "ERF.PRECISE": "ERF.PRECISE(x)",
  "ERFC.PRECISE": "ERFC.PRECISE(x)",
  DEC2HEX: "DEC2HEX(number, [places])",
  HEX2DEC: "HEX2DEC(number)",
  DSUM: "DSUM(database, field, criteria)",
  DCOUNT: "DCOUNT(database, field, criteria)",
  DCOUNTA: "DCOUNTA(database, field, criteria)",
  DAVERAGE: "DAVERAGE(database, field, criteria)",
  DMAX: "DMAX(database, field, criteria)",
  DMIN: "DMIN(database, field, criteria)",
  DPRODUCT: "DPRODUCT(database, field, criteria)",
  DSTDEV: "DSTDEV(database, field, criteria)",
  DSTDEVP: "DSTDEVP(database, field, criteria)",
  DVAR: "DVAR(database, field, criteria)",
  DVARP: "DVARP(database, field, criteria)",
};



const FORMULA_ENGINE_ARGUMENT_OVERRIDES = {
  VLOOKUP: ["lookup_value", "table_array", "col_index_num", "range_lookup"],
  HLOOKUP: ["lookup_value", "table_array", "row_index_num", "range_lookup"],
  MATCH: ["lookup_value", "lookup_array", "match_type"],
  INDEX: ["array", "row_num", "column_num"],
  XLOOKUP: ["lookup_value", "lookup_array", "return_array", "if_not_found", "match_mode", "search_mode"],
  IFS: ["logical_test1", "value_if_true1"],
  XOR: ["logical1"],
  LET: ["name1", "name_value1", "calculation_or_name2", "name_value2"],
  LAMBDA: ["parameter1", "calculation"],
  MAKEARRAY: ["rows", "cols", "lambda"],
  BYROW: ["array", "lambda"],
  BYCOL: ["array", "lambda"],
  MAP: ["array1", "lambda_or_array2"],
  REDUCE: ["initial_value", "array", "lambda"],
  SCAN: ["initial_value", "array", "lambda"],
  COUNTIF: ["range", "criteria"],
  COUNTIFS: ["criteria_range1", "criteria1"],
  SUMIF: ["range", "criteria", "sum_range"],
  SUMIFS: ["sum_range", "criteria_range1", "criteria1"],
  AVERAGEIF: ["range", "criteria", "average_range"],
  AVERAGEIFS: ["average_range", "criteria_range1", "criteria1"],
  MAXIFS: ["max_range", "criteria_range1", "criteria1"],
  MINIFS: ["min_range", "criteria_range1", "criteria1"],
  LOOKUP: ["lookup_value", "lookup_vector", "result_vector"],
  FILTER: ["array", "include", "if_empty"],
  SORT: ["array", "sort_index", "sort_order", "by_col"],
  SORTBY: ["array", "by_array1", "sort_order1"],
  UNIQUE: ["array", "by_col", "exactly_once"],
  SEQUENCE: ["rows", "columns", "start", "step"],
  TAKE: ["array", "rows", "columns"],
  DROP: ["array", "rows", "columns"],
  VSTACK: ["array1", "array2"],
  HSTACK: ["array1", "array2"],
  CHOOSEROWS: ["array", "row_num1", "row_num2"],
  CHOOSECOLS: ["array", "col_num1", "col_num2"],
  TOCOL: ["array", "ignore", "scan_by_column"],
  TOROW: ["array", "ignore", "scan_by_column"],
  WRAPROWS: ["vector", "wrap_count", "pad_with"],
  WRAPCOLS: ["vector", "wrap_count", "pad_with"],
  EXPAND: ["array", "rows", "columns", "pad_with"],
  OFFSET: ["reference", "rows", "cols", "height", "width"],
  ADDRESS: ["row_num", "column_num", "abs_num", "a1", "sheet_text"],
  SUBTOTAL: ["function_num", "ref1"],
  AGGREGATE: ["function_num", "options", "ref1", "ref2"],
  PMT: ["rate", "nper", "pv", "fv", "type"],
  PV: ["rate", "nper", "pmt", "fv", "type"],
  FV: ["rate", "nper", "pmt", "pv", "type"],
  NPV: ["rate", "value1"],
  FVSCHEDULE: ["principal", "schedule"],
  XNPV: ["rate", "values", "dates"],
  XIRR: ["values", "dates", "guess"],
  IRR: ["values", "guess"],
  MIRR: ["values", "finance_rate", "reinvest_rate"],
  RATE: ["nper", "pmt", "pv", "fv", "type", "guess"],
  NPER: ["rate", "pmt", "pv", "fv", "type"],
  DB: ["cost", "salvage", "life", "period", "month"],
  DDB: ["cost", "salvage", "life", "period", "factor"],
  VDB: ["cost", "salvage", "life", "start_period", "end_period", "factor", "no_switch"],
  AMORLINC: ["cost", "date_purchased", "first_period", "salvage", "period", "rate", "basis"],
  AMORDEGRC: ["cost", "date_purchased", "first_period", "salvage", "period", "rate", "basis"],
  SLN: ["cost", "salvage", "life"],
  SYD: ["cost", "salvage", "life", "per"],
  DOLLARDE: ["fractional_dollar", "fraction"],
  DOLLARFR: ["decimal_dollar", "fraction"],
  RRI: ["nper", "pv", "fv"],
  PDURATION: ["rate", "pv", "fv"],
  ISPMT: ["rate", "per", "nper", "pv"],
  TBILLEQ: ["settlement", "maturity", "discount"],
  TBILLPRICE: ["settlement", "maturity", "discount"],
  TBILLYIELD: ["settlement", "maturity", "pr"],
  DISC: ["settlement", "maturity", "pr", "redemption", "basis"],
  PRICEDISC: ["settlement", "maturity", "discount", "redemption", "basis"],
  YIELDDISC: ["settlement", "maturity", "pr", "redemption", "basis"],
  INTRATE: ["settlement", "maturity", "investment", "redemption", "basis"],
  RECEIVED: ["settlement", "maturity", "investment", "discount", "basis"],
  ACCRINT: ["issue", "first_interest", "settlement", "rate", "par", "frequency", "basis", "calc_method"],
  ACCRINTM: ["issue", "settlement", "rate", "par", "basis"],
  PRICEMAT: ["settlement", "maturity", "issue", "rate", "yld", "basis"],
  YIELDMAT: ["settlement", "maturity", "issue", "rate", "pr", "basis"],
  COUPDAYBS: ["settlement", "maturity", "frequency", "basis"],
  COUPDAYS: ["settlement", "maturity", "frequency", "basis"],
  COUPDAYSNC: ["settlement", "maturity", "frequency", "basis"],
  COUPNCD: ["settlement", "maturity", "frequency", "basis"],
  COUPPCD: ["settlement", "maturity", "frequency", "basis"],
  COUPNUM: ["settlement", "maturity", "frequency", "basis"],
  PRICE: ["settlement", "maturity", "rate", "yld", "redemption", "frequency", "basis"],
  YIELD: ["settlement", "maturity", "rate", "pr", "redemption", "frequency", "basis"],
  DURATION: ["settlement", "maturity", "coupon", "yld", "frequency", "basis"],
  MDURATION: ["settlement", "maturity", "coupon", "yld", "frequency", "basis"],
  IPMT: ["rate", "per", "nper", "pv", "fv", "type"],
  PPMT: ["rate", "per", "nper", "pv", "fv", "type"],
  FORMULATEXT: ["reference"],
  ISFORMULA: ["reference"],
  ISREF: ["value"],
  SHEET: ["value"],
  SHEETS: ["reference"],
  CELL: ["info_type", "reference"],
  INFO: ["type_text"],
  ROW: ["reference"],
  COLUMN: ["reference"],
  MMULT: ["array1", "array2"],
  MDETERM: ["array"],
  MINVERSE: ["array"],
  MUNIT: ["dimension"],
  HYPERLINK: ["link_location", "friendly_name"],
  T: ["value"],
  DATE: ["year", "month", "day"],
  TIME: ["hour", "minute", "second"],
  EDATE: ["start_date", "months"],
  EOMONTH: ["start_date", "months"],
  DATEDIF: ["start_date", "end_date", "unit"],
  ISOWEEKNUM: ["date"],
  NETWORKDAYS: ["start_date", "end_date", "holidays"],
  "NETWORKDAYS.INTL": ["start_date", "end_date", "weekend", "holidays"],
  WORKDAY: ["start_date", "days", "holidays"],
  "WORKDAY.INTL": ["start_date", "days", "weekend", "holidays"],
  YEARFRAC: ["start_date", "end_date", "basis"],
  TEXT: ["value", "format_text"],
  TEXTJOIN: ["delimiter", "ignore_empty", "text1"],
  TEXTSPLIT: ["text", "col_delimiter", "row_delimiter", "ignore_empty", "match_mode", "pad_with"],
  TEXTBEFORE: ["text", "delimiter", "instance_num", "match_mode", "match_end", "if_not_found"],
  TEXTAFTER: ["text", "delimiter", "instance_num", "match_mode", "match_end", "if_not_found"],
  CONCATENATE: ["text1", "text2"],
  LEFT: ["text", "num_chars"],
  RIGHT: ["text", "num_chars"],
  MID: ["text", "start_num", "num_chars"],
  FIND: ["find_text", "within_text", "start_num"],
  SEARCH: ["find_text", "within_text", "start_num"],
  SUBSTITUTE: ["text", "old_text", "new_text", "instance_num"],
  REPLACE: ["old_text", "start_num", "num_chars", "new_text"],
  NUMBERVALUE: ["text", "decimal_separator", "group_separator"],
  T_TEST: ["array1", "array2", "tails", "type"],
  "T.TEST": ["array1", "array2", "tails", "type"],
  "CHISQ.TEST": ["actual_range", "expected_range"],
  "F.TEST": ["array1", "array2"],
  "Z.TEST": ["array", "x", "sigma"],
  KURT: ["number1", "number2"],
  MODE: ["number1", "number2"],
  "MODE.MULT": ["number1", "number2"],
  RANK: ["number", "ref", "order"],
  "RANK.AVG": ["number", "ref", "order"],
  "PERCENTILE.INC": ["array", "k"],
  PERCENTILE: ["array", "k"],
  "PERCENTILE.EXC": ["array", "k"],
  QUARTILE: ["array", "quart"],
  "QUARTILE.INC": ["array", "quart"],
  "QUARTILE.EXC": ["array", "quart"],
  "PERCENTRANK.INC": ["array", "x", "significance"],
  "PERCENTRANK.EXC": ["array", "x", "significance"],
  PROB: ["x_range", "prob_range", "lower_limit", "upper_limit"],
  FREQUENCY: ["data_array", "bins_array"],
  PEARSON: ["array1", "array2"],
  RSQ: ["known_y's", "known_x's"],
  STEYX: ["known_y's", "known_x's"],
  SKEW: ["number1", "number2"],
  "SKEW.P": ["number1", "number2"],
  LINEST: ["known_y's", "known_x's", "const", "stats"],
  LOGEST: ["known_y's", "known_x's", "const", "stats"],
  FORECAST: ["x", "known_y's", "known_x's"],
  "FORECAST.LINEAR": ["x", "known_y's", "known_x's"],
  TREND: ["known_y's", "known_x's", "new_x's", "const"],
  GROWTH: ["known_y's", "known_x's", "new_x's", "const"],
  GAMMA: ["number"],
  "GAMMA.DIST": ["x", "alpha", "beta", "cumulative"],
  "GAMMA.INV": ["probability", "alpha", "beta"],
  GAMMALN: ["x"],
  "GAMMALN.PRECISE": ["x"],
  "BETA.DIST": ["x", "alpha", "beta", "cumulative", "A", "B"],
  "BETA.INV": ["probability", "alpha", "beta", "A", "B"],
  "BINOM.DIST.RANGE": ["trials", "probability_s", "number_s", "number_s2"],
  TRIMMEAN: ["array", "percent"],
  "CHISQ.DIST": ["x", "deg_freedom", "cumulative"],
  "CHISQ.INV": ["probability", "deg_freedom"],
  "WEIBULL.DIST": ["x", "alpha", "beta", "cumulative"],
  "HYPGEOM.DIST": ["sample_s", "number_sample", "population_s", "number_pop", "cumulative"],
  "CONFIDENCE.NORM": ["alpha", "standard_dev", "size"],
  "CONFIDENCE.T": ["alpha", "standard_dev", "size"],
  "F.DIST": ["x", "deg_freedom1", "deg_freedom2", "cumulative"],
  "F.DIST.RT": ["x", "deg_freedom1", "deg_freedom2"],
  "F.INV": ["probability", "deg_freedom1", "deg_freedom2"],
  "F.INV.RT": ["probability", "deg_freedom1", "deg_freedom2"],
  "T.DIST": ["x", "deg_freedom", "cumulative"],
  "T.DIST.RT": ["x", "deg_freedom"],
  "T.DIST.2T": ["x", "deg_freedom"],
  "T.INV": ["probability", "deg_freedom"],
  "T.INV.2T": ["probability", "deg_freedom"],
  "LOGNORM.DIST": ["x", "mean", "standard_dev", "cumulative"],
  "LOGNORM.INV": ["probability", "mean", "standard_dev"],
  "NEGBINOM.DIST": ["number_f", "number_s", "probability_s", "cumulative"],
  ERF: ["lower_limit", "upper_limit"],
  ERFC: ["x"],
  "ERF.PRECISE": ["x"],
  PHI: ["x"],
  GAUSS: ["z"],
  BITAND: ["number1", "number2"],
  BITOR: ["number1", "number2"],
  BITXOR: ["number1", "number2"],
  BITLSHIFT: ["number", "shift_amount"],
  BITRSHIFT: ["number", "shift_amount"],
  DEC2BIN: ["number", "places"],
  DEC2OCT: ["number", "places"],
  DEC2HEX: ["number", "places"],
  BIN2DEC: ["number"],
  BIN2OCT: ["number", "places"],
  BIN2HEX: ["number", "places"],
  OCT2BIN: ["number", "places"],
  OCT2DEC: ["number"],
  OCT2HEX: ["number", "places"],
  HEX2BIN: ["number", "places"],
  HEX2OCT: ["number", "places"],
  HEX2DEC: ["number"],
  DELTA: ["number1", "number2"],
  GESTEP: ["number", "step"],
  SQRTPI: ["number"],
  SUMX2MY2: ["array_x", "array_y"],
  SUMX2PY2: ["array_x", "array_y"],
  SUMXMY2: ["array_x", "array_y"],
  COMPLEX: ["real_num", "i_num", "suffix"],
  IMARGUMENT: ["inumber"],
  IMEXP: ["inumber"],
  IMLN: ["inumber"],
  IMLOG10: ["inumber"],
  IMLOG2: ["inumber"],
  IMPOWER: ["inumber", "number"],
  IMSQRT: ["inumber"],
  IMSIN: ["inumber"],
  IMCOS: ["inumber"],
  IMTAN: ["inumber"],
  IMSEC: ["inumber"],
  IMCSC: ["inumber"],
  IMCOT: ["inumber"],
  IMSINH: ["inumber"],
  IMCOSH: ["inumber"],
  IMSECH: ["inumber"],
  IMCSCH: ["inumber"],
  BESSELI: ["x", "n"],
  BESSELJ: ["x", "n"],
  BESSELK: ["x", "n"],
  BESSELY: ["x", "n"],
  "ERFC.PRECISE": ["x"],
  BETADIST: ["x", "alpha", "beta", "A", "B"],
  BETAINV: ["probability", "alpha", "beta", "A", "B"],
  BINOMDIST: ["number_s", "trials", "probability_s", "cumulative"],
  CHIDIST: ["x", "deg_freedom"],
  CHIDISTRT: ["x", "deg_freedom"],
  CHIINV: ["probability", "deg_freedom"],
  CHIINVRT: ["probability", "deg_freedom"],
  CHITEST: ["actual_range", "expected_range"],
  CONFIDENCE: ["alpha", "standard_dev", "size"],
  COVAR: ["array1", "array2"],
  COVARIANCEP: ["array1", "array2"],
  COVARIANCES: ["array1", "array2"],
  CRITBINOM: ["trials", "probability_s", "alpha"],
  EXPONDIST: ["x", "lambda", "cumulative"],
  FDIST: ["x", "deg_freedom1", "deg_freedom2"],
  FDISTRT: ["x", "deg_freedom1", "deg_freedom2"],
  FINV: ["probability", "deg_freedom1", "deg_freedom2"],
  FINVRT: ["probability", "deg_freedom1", "deg_freedom2"],
  FTEST: ["array1", "array2"],
  GAMMADIST: ["x", "alpha", "beta", "cumulative"],
  GAMMAINV: ["probability", "alpha", "beta"],
  HYPGEOMDIST: ["sample_s", "number_sample", "population_s", "number_pop"],
  LOGINV: ["probability", "mean", "standard_dev"],
  LOGNORMDIST: ["x", "mean", "standard_dev"],
  LOGNORMINV: ["probability", "mean", "standard_dev"],
  NEGBINOMDIST: ["number_f", "number_s", "probability_s"],
  NORMDIST: ["x", "mean", "standard_dev", "cumulative"],
  NORMINV: ["probability", "mean", "standard_dev"],
  NORMSDIST: ["z"],
  NORMSINV: ["probability"],
  POISSON: ["x", "mean", "cumulative"],
  POISSONDIST: ["x", "mean", "cumulative"],
  STDEV: ["number1", "number2"],
  STDEVP: ["number1", "number2"],
  TINV: ["probability", "deg_freedom"],
  TINV2T: ["probability", "deg_freedom"],
  TDIST: ["x", "deg_freedom", "tails"],
  TDIST2T: ["x", "deg_freedom"],
  TDISTRT: ["x", "deg_freedom"],
  TTEST: ["array1", "array2", "tails", "type"],
  VAR: ["number1", "number2"],
  VARP: ["number1", "number2"],
  WEIBULL: ["x", "alpha", "beta", "cumulative"],
  WEIBULLDIST: ["x", "alpha", "beta", "cumulative"],
  ZTEST: ["array", "x", "sigma"],
  DSUM: ["database", "field", "criteria"],
  DCOUNT: ["database", "field", "criteria"],
  DCOUNTA: ["database", "field", "criteria"],
  DAVERAGE: ["database", "field", "criteria"],
  DMAX: ["database", "field", "criteria"],
  DMIN: ["database", "field", "criteria"],
  DPRODUCT: ["database", "field", "criteria"],
  DSTDEV: ["database", "field", "criteria"],
  DSTDEVP: ["database", "field", "criteria"],
  DVAR: ["database", "field", "criteria"],
  DVARP: ["database", "field", "criteria"],
};



const INSERT_FUNCTION_RECENT_NAMES = ["SUM", "AVERAGE", "COUNT", "IF", "XLOOKUP", "VLOOKUP", "SUMIF", "COUNTIF", "TEXT", "TODAY"];


const MAX_RECENT_FORMULA_FUNCTIONS = 10;


const INSERT_FUNCTION_CATEGORY_ORDER = [
  "最近使用した関数",
  "すべて表示",
  "財務",
  "日付/時刻",
  "数学/三角",
  "統計",
  "検索/行列",
  "データベース",
  "文字列",
  "論理",
  "論理/情報",
  "情報",
  "エンジニアリング",
  "互換性",
];



const FORMULA_RIBBON_FUNCTION_MENUS = {
  recent: {
    label: "最近使用した関数",
    category: "最近使用した関数",
    functions: [],
  },
  financial: {
    label: "財務",
    category: "財務",
    functions: ["PMT", "PV", "FV", "FVSCHEDULE", "NPV", "XNPV", "XIRR", "IRR", "MIRR", "RATE", "NPER", "IPMT", "PPMT", "CUMIPMT", "CUMPRINC", "DB", "DDB", "VDB", "AMORLINC", "AMORDEGRC", "SLN", "SYD", "DOLLARDE", "DOLLARFR", "EFFECT", "NOMINAL", "RRI", "PDURATION", "ISPMT", "DISC", "PRICEDISC", "YIELDDISC", "INTRATE", "RECEIVED", "ACCRINT", "ACCRINTM", "PRICEMAT", "YIELDMAT", "COUPDAYBS", "COUPDAYS", "COUPDAYSNC", "COUPNCD", "COUPPCD", "COUPNUM", "PRICE", "YIELD", "DURATION", "MDURATION", "TBILLEQ", "TBILLPRICE", "TBILLYIELD"],
  },
  math: {
    label: "数学/三角",
    category: "数学/三角",
    functions: ["SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "SUMSQ", "SUMX2MY2", "SUMX2PY2", "SUMXMY2", "SUBTOTAL", "AGGREGATE", "PRODUCT", "ROUND", "ROUNDUP", "ROUNDDOWN", "INT", "ABS", "MOD", "POWER", "SQRT", "SQRTPI", "RAND", "RANDBETWEEN"],
  },
  logical: {
    label: "論理",
    category: "論理",
    functions: ["IF", "IFS", "LET", "AND", "OR", "NOT", "XOR", "TRUE", "FALSE", "IFERROR", "IFNA", "SWITCH"],
  },
  lookup: {
    label: "検索/行列",
    category: "検索/行列",
    functions: ["XLOOKUP", "VLOOKUP", "HLOOKUP", "LOOKUP", "INDEX", "MATCH", "XMATCH", "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "TAKE", "DROP", "VSTACK", "HSTACK", "MAKEARRAY", "BYROW", "BYCOL", "MAP", "REDUCE", "SCAN", "CHOOSEROWS", "CHOOSECOLS", "TOCOL", "TOROW", "WRAPROWS", "WRAPCOLS", "EXPAND", "MMULT", "MDETERM", "MINVERSE", "MUNIT", "OFFSET", "INDIRECT", "ADDRESS", "ROW", "ROWS", "COLUMN", "COLUMNS", "SHEET", "SHEETS", "CHOOSE", "TRANSPOSE", "HYPERLINK"],
  },
  date: {
    label: "日付/時刻",
    category: "日付/時刻",
    functions: ["TODAY", "NOW", "DATE", "TIME", "YEAR", "MONTH", "DAY", "HOUR", "MINUTE", "SECOND", "WEEKDAY", "WEEKNUM", "ISOWEEKNUM", "EOMONTH", "EDATE", "DATEVALUE", "TIMEVALUE", "NETWORKDAYS", "NETWORKDAYS.INTL", "WORKDAY", "WORKDAY.INTL", "DAYS", "DAYS360", "YEARFRAC"],
  },
  text: {
    label: "文字列",
    category: "文字列",
    functions: ["TEXTJOIN", "TEXTSPLIT", "TEXTBEFORE", "TEXTAFTER", "CONCAT", "CONCATENATE", "TEXT", "VALUE", "LEFT", "RIGHT", "MID", "LEN", "FIND", "SEARCH", "SUBSTITUTE", "REPLACE", "TRIM", "UPPER", "LOWER", "PROPER", "EXACT"],
  },
  more: {
    label: "その他の関数",
    category: "すべて表示",
    functions: ["COUNTBLANK", "COUNTA", "STDEV.S", "STDEV.P", "VAR.S", "VAR.P", "MEDIAN", "MODE.SNGL", "MODE.MULT", "KURT", "SKEW", "SKEW.P", "LARGE", "SMALL", "RANK.EQ", "RANK.AVG", "PERCENTILE.INC", "PERCENTRANK.INC", "PERCENTRANK.EXC", "PROB", "FREQUENCY", "TRIMMEAN", "LINEST", "LOGEST", "FORECAST.LINEAR", "TREND", "GROWTH", "CORREL", "PEARSON", "RSQ", "STEYX", "COVARIANCE.S", "GAMMALN.PRECISE", "IMABS", "IMSUM", "IMPRODUCT", "BESSELI", "BESSELJ", "BESSELK", "BESSELY", "ERF.PRECISE", "ERFC.PRECISE", "FORMULATEXT", "ISFORMULA", "ISREF", "CELL", "INFO", "ISBLANK", "ISERROR", "ISNUMBER", "NORMDIST", "NORMSDIST", "BETADIST", "BINOMDIST", "BINOM.DIST.RANGE", "TTEST"],
  },
};



const MAX_FUNCTION_ARGUMENT_ROWS = 16;



const FORMULA_ARGUMENT_DEFAULT_OVERRIDES = {
  IF: ["1=1", "1", "0"],
  IFS: ["1=1", "1"],
  LET: ["x", "1", "x+1"],
  LAMBDA: ["x", "x+1"],
  MAKEARRAY: ["2", "2", "LAMBDA(r,c,r+c)"],
  BYROW: [formulaTemplateReference, "LAMBDA(row,SUM(row))"],
  BYCOL: [formulaTemplateReference, "LAMBDA(col,SUM(col))"],
  MAP: [formulaTemplateReference, "LAMBDA(x,x)"],
  REDUCE: ["0", formulaTemplateReference, "LAMBDA(acc,value,acc+value)"],
  SCAN: ["0", formulaTemplateReference, "LAMBDA(acc,value,acc+value)"],
  SORTBY: [formulaTemplateReference, formulaTemplateReference, "-1"],
  AND: ["1=1"],
  OR: ["1=1"],
  XOR: ["1=1"],
  NOT: ["1=1"],
  IFERROR: [formulaTemplateReference, "0"],
  SUM: [formulaTemplateReference],
  AVERAGE: [formulaTemplateReference],
  COUNT: [formulaTemplateReference],
  MAX: [formulaTemplateReference],
  MIN: [formulaTemplateReference],
  SUMPRODUCT: [formulaTemplateReference],
  COUNTIF: [formulaTemplateReference, '"<>"'],
  COUNTIFS: [formulaTemplateReference, '"<>"'],
  SUMIF: [formulaTemplateReference, '"<>"'],
  SUMIFS: [formulaTemplateReference, formulaTemplateReference, '"<>"'],
  AVERAGEIF: [formulaTemplateReference, '"<>"'],
  AVERAGEIFS: [formulaTemplateReference, formulaTemplateReference, '"<>"'],
  MAXIFS: [formulaTemplateReference, formulaTemplateReference, '"<>"'],
  MINIFS: [formulaTemplateReference, formulaTemplateReference, '"<>"'],
  AGGREGATE: ["9", "0", formulaTemplateReference],
  TEXTJOIN: ['","', "1=1", formulaTemplateReference],
  TEXTSPLIT: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), '","'],
  TEXTBEFORE: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), '","'],
  TEXTAFTER: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), '","'],
  XLOOKUP: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), formulaTemplateReference, formulaTemplateReference],
  VLOOKUP: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), formulaTemplateReference, "1", "TRUE"],
  HLOOKUP: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), formulaTemplateReference, "1", "TRUE"],
  LOOKUP: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), formulaTemplateReference],
  INDEX: [formulaTemplateReference, "1"],
  MATCH: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), formulaTemplateReference, "0"],
  FILTER: [formulaTemplateReference, "TRUE"],
  SORT: [formulaTemplateReference],
  UNIQUE: [formulaTemplateReference],
  SEQUENCE: ["3", "1", "1", "1"],
  TAKE: [formulaTemplateReference, "1"],
  DROP: [formulaTemplateReference, "1"],
  VSTACK: [formulaTemplateReference, formulaTemplateReference],
  HSTACK: [formulaTemplateReference, formulaTemplateReference],
  CHOOSEROWS: [formulaTemplateReference, "1"],
  CHOOSECOLS: [formulaTemplateReference, "1"],
  TOCOL: [formulaTemplateReference],
  TOROW: [formulaTemplateReference],
  WRAPROWS: [formulaTemplateReference, "2"],
  WRAPCOLS: [formulaTemplateReference, "2"],
  EXPAND: [formulaTemplateReference, "3", "3"],
  OFFSET: [formulaTemplateReference, "1", "0"],
  ROW: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  COLUMN: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  ROWS: [formulaTemplateReference],
  COLUMNS: [formulaTemplateReference],
  HYPERLINK: ['"https://example.com"', '"リンク"'],
  DATE: ["2000", "1", "1"],
  TIME: ["12", "0", "0"],
  DATEDIF: ["DATE(2000,1,1)", "TODAY()", '"Y"'],
  ISOWEEKNUM: ["TODAY()"],
  EDATE: ["TODAY()", "1"],
  EOMONTH: ["TODAY()", "1"],
  NETWORKDAYS: ["TODAY()", "TODAY()+7"],
  "NETWORKDAYS.INTL": ["TODAY()", "TODAY()+7"],
  WORKDAY: ["TODAY()", "1"],
  "WORKDAY.INTL": ["TODAY()", "1"],
  DAYS: ["TODAY()+1", "TODAY()"],
  DAYS360: ["TODAY()", "TODAY()+30"],
  YEARFRAC: ["TODAY()", "TODAY()+365"],
  LEFT: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), "1"],
  RIGHT: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), "1"],
  MID: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), "1", "1"],
  FIND: ['""', () => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  SEARCH: ['""', () => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  SUBSTITUTE: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), '""', '""'],
  REPLACE: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), "1", "1", '""'],
  TEXT: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), '"0"'],
  VALUE: ['"1"'],
  NUMBERVALUE: ['"1"'],
  ROUND: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), "0"],
  ROUNDUP: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), "0"],
  ROUNDDOWN: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), "0"],
  KURT: [formulaTemplateReference],
  MODE: [formulaTemplateReference],
  "MODE.MULT": [formulaTemplateReference],
  RANK: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), formulaTemplateReference],
  "RANK.AVG": [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), formulaTemplateReference],
  PERCENTILE: [formulaTemplateReference, "0.5"],
  "PERCENTILE.INC": [formulaTemplateReference, "0.5"],
  "PERCENTRANK.INC": [formulaTemplateReference, () => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  "PERCENTRANK.EXC": [formulaTemplateReference, () => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  PROB: [formulaTemplateReference, formulaTemplateReference, () => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  QUARTILE: [formulaTemplateReference, "1"],
  FREQUENCY: [formulaTemplateReference, formulaTemplateReference],
  LINEST: [formulaTemplateReference, formulaTemplateReference, "TRUE", "FALSE"],
  LOGEST: [formulaTemplateReference, formulaTemplateReference, "TRUE", "FALSE"],
  FORECAST: ["1", formulaTemplateReference, formulaTemplateReference],
  "FORECAST.LINEAR": ["1", formulaTemplateReference, formulaTemplateReference],
  TREND: [formulaTemplateReference, formulaTemplateReference, formulaTemplateReference, "TRUE"],
  GROWTH: [formulaTemplateReference, formulaTemplateReference, formulaTemplateReference, "TRUE"],
  MOD: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), "1"],
  CEILING: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), "1"],
  FLOOR: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), "1"],
  MROUND: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), "1"],
  QUOTIENT: ["1", "1"],
  COMBIN: ["1", "1"],
  PMT: ["0.05/12", "12", "1000"],
  PV: ["0.05/12", "12", "-100"],
  FV: ["0.05/12", "12", "-100"],
  RATE: ["12", "-100", "1000"],
  NPER: ["0.05/12", "-100", "1000"],
  IPMT: ["0.05/12", "1", "12", "1000"],
  PPMT: ["0.05/12", "1", "12", "1000"],
  NPV: ["0.05", formulaTemplateReference],
  IRR: [formulaTemplateReference],
  MIRR: [formulaTemplateReference, "0.05", "0.05"],
  DB: ["1000", "100", "5", "1"],
  DDB: ["1000", "100", "5", "1"],
  VDB: ["1000", "100", "5", "1", "2"],
  AMORLINC: ["1000", "DATE(2023,1,1)", "DATE(2023,7,1)", "100", "1", "0.15"],
  AMORDEGRC: ["1000", "DATE(2023,1,1)", "DATE(2023,7,1)", "100", "1", "0.15"],
  SLN: ["1000", "100", "5"],
  SYD: ["1000", "100", "5", "1"],
  FVSCHEDULE: ["1000", formulaTemplateReference],
  CUMIPMT: ["0.05/12", "12", "1000", "1", "12", "0"],
  CUMPRINC: ["0.05/12", "12", "1000", "1", "12", "0"],
  EFFECT: ["0.05", "12"],
  NOMINAL: ["0.05", "12"],
  XNPV: ["0.05", formulaTemplateReference, formulaTemplateReference],
  XIRR: [formulaTemplateReference, formulaTemplateReference],
  DOLLARDE: ["1.02", "16"],
  DOLLARFR: ["1.125", "16"],
  RRI: ["12", "1000", "2000"],
  PDURATION: ["0.05", "1000", "2000"],
  ISPMT: ["0.05/12", "1", "12", "1000"],
  TBILLEQ: ["TODAY()", "TODAY()+90", "0.05"],
  TBILLPRICE: ["TODAY()", "TODAY()+90", "0.05"],
  TBILLYIELD: ["TODAY()", "TODAY()+90", "98"],
  DISC: ["TODAY()", "TODAY()+180", "97", "100"],
  PRICEDISC: ["TODAY()", "TODAY()+180", "0.05", "100"],
  YIELDDISC: ["TODAY()", "TODAY()+180", "97", "100"],
  INTRATE: ["TODAY()", "TODAY()+180", "970", "1000"],
  RECEIVED: ["TODAY()", "TODAY()+180", "970", "0.05"],
  ACCRINT: ["TODAY()", "TODAY()+180", "TODAY()+90", "0.05", "1000", "2"],
  ACCRINTM: ["TODAY()", "TODAY()+180", "0.05", "1000"],
  PRICEMAT: ["TODAY()+30", "TODAY()+180", "TODAY()", "0.05", "0.04"],
  YIELDMAT: ["TODAY()+30", "TODAY()+180", "TODAY()", "0.05", "98"],
  COUPDAYBS: ["TODAY()", "TODAY()+365", "2"],
  COUPDAYS: ["TODAY()", "TODAY()+365", "2"],
  COUPDAYSNC: ["TODAY()", "TODAY()+365", "2"],
  COUPNCD: ["TODAY()", "TODAY()+365", "2"],
  COUPPCD: ["TODAY()", "TODAY()+365", "2"],
  COUPNUM: ["TODAY()", "TODAY()+365", "2"],
  PRICE: ["TODAY()", "TODAY()+365", "0.05", "0.04", "100", "2"],
  YIELD: ["TODAY()", "TODAY()+365", "0.05", "98", "100", "2"],
  DURATION: ["TODAY()", "TODAY()+365", "0.05", "0.04", "2"],
  MDURATION: ["TODAY()", "TODAY()+365", "0.05", "0.04", "2"],
  FORMULATEXT: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  ISFORMULA: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  ISREF: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  SHEET: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  SHEETS: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  CELL: ['"address"', () => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  INFO: ['"recalc"'],
  MMULT: [formulaTemplateReference, formulaTemplateReference],
  MDETERM: [formulaTemplateReference],
  MINVERSE: [formulaTemplateReference],
  MUNIT: ["2"],
  ISERR: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  ISLOGICAL: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  ISNONTEXT: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  ISEVEN: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  ISODD: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  LOG10: ["10"],
  SQRTPI: ["1"],
  FACTDOUBLE: ["5"],
  COMBINA: ["3", "2"],
  PERMUT: ["5", "2"],
  PERMUTATIONA: ["3", "2"],
  GCD: [formulaTemplateReference],
  LCM: [formulaTemplateReference],
  MULTINOMIAL: [formulaTemplateReference],
  SIN: ["0"],
  COS: ["0"],
  TAN: ["0.5"],
  ASIN: ["0.5"],
  ACOS: ["0.5"],
  ATAN: ["1"],
  DEGREES: ["1"],
  RADIANS: ["180"],
  SERIESSUM: ["1", "0", "1", formulaTemplateReference],
  SUMSQ: [formulaTemplateReference],
  SUMX2MY2: [formulaTemplateReference, formulaTemplateReference],
  SUMX2PY2: [formulaTemplateReference, formulaTemplateReference],
  SUMXMY2: [formulaTemplateReference, formulaTemplateReference],
  EVEN: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  ODD: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  "ISO.CEILING": [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), "1"],
  ATAN2: ["1", "1"],
  SINH: ["1"],
  COSH: ["1"],
  TANH: ["1"],
  ASINH: ["1"],
  ACOSH: ["1"],
  ATANH: ["0.5"],
  COT: ["1"],
  SEC: ["1"],
  CSC: ["1"],
  ACOT: ["1"],
  COTH: ["1"],
  SECH: ["1"],
  CSCH: ["1"],
  AVEDEV: [formulaTemplateReference],
  DEVSQ: [formulaTemplateReference],
  GEOMEAN: [formulaTemplateReference],
  HARMEAN: [formulaTemplateReference],
  RSQ: [formulaTemplateReference, formulaTemplateReference],
  PEARSON: [formulaTemplateReference, formulaTemplateReference],
  "Z.TEST": [formulaTemplateReference, () => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  "F.TEST": [formulaTemplateReference, formulaTemplateReference],
  STEYX: [formulaTemplateReference, formulaTemplateReference],
  SKEW: [formulaTemplateReference],
  "SKEW.P": [formulaTemplateReference],
  "EXPON.DIST": ["1", "1", "TRUE"],
  FISHER: ["0.5"],
  FISHERINV: ["0.5"],
  GAMMA: ["1"],
  "GAMMA.DIST": ["1", "1", "1", "TRUE"],
  "GAMMA.INV": ["0.5", "1", "1"],
  GAMMALN: ["1"],
  "GAMMALN.PRECISE": ["1"],
  "BETA.DIST": ["0.5", "1", "1", "TRUE"],
  "BETA.INV": ["0.5", "1", "1"],
  "BINOM.DIST": ["1", "10", "0.5", "TRUE"],
  "BINOM.DIST.RANGE": ["10", "0.5", "3", "5"],
  "BINOM.INV": ["10", "0.5", "0.5"],
  "CHISQ.DIST": ["1", "1", "TRUE"],
  "CHISQ.INV": ["0.5", "1"],
  "F.DIST": ["1", "1", "1", "TRUE"],
  "F.DIST.RT": ["1", "1", "1"],
  "F.INV": ["0.5", "1", "1"],
  "F.INV.RT": ["0.5", "1", "1"],
  "T.DIST": ["1", "1", "TRUE"],
  "T.DIST.RT": ["1", "1"],
  "T.DIST.2T": ["1", "1"],
  "T.INV": ["0.5", "1"],
  "T.INV.2T": ["0.5", "1"],
  "LOGNORM.DIST": ["1", "0", "1", "TRUE"],
  "LOGNORM.INV": ["0.5", "0", "1"],
  "NEGBINOM.DIST": ["1", "1", "0.5", "TRUE"],
  ERF: ["1"],
  ERFC: ["1"],
  "ERF.PRECISE": ["1"],
  PHI: ["1"],
  GAUSS: ["1"],
  "NORM.S.DIST": ["1", "TRUE"],
  "NORM.S.INV": ["0.5"],
  STANDARDIZE: ["1", "0", "1"],
  TRIMMEAN: [formulaTemplateReference, "0.2"],
  "POISSON.DIST": ["1", "1", "TRUE"],
  "WEIBULL.DIST": ["1", "1", "1", "TRUE"],
  "HYPGEOM.DIST": ["1", "10", "5", "20", "TRUE"],
  "CONFIDENCE.NORM": ["0.05", "1", "30"],
  "CONFIDENCE.T": ["0.05", "1", "30"],
  SPLIT: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference()), '","'],
  T: [() => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  UNICODE: ['"A"'],
  UNICHAR: ["65"],
  BITAND: ["1", "1"],
  BITOR: ["1", "1"],
  BITXOR: ["1", "1"],
  BITLSHIFT: ["1", "1"],
  BITRSHIFT: ["2", "1"],
  BIN2DEC: ['"10"'],
  BIN2HEX: ['"10"'],
  BIN2OCT: ['"10"'],
  OCT2BIN: ['"10"'],
  OCT2DEC: ['"10"'],
  OCT2HEX: ['"10"'],
  HEX2BIN: ['"10"'],
  HEX2OCT: ['"10"'],
  COMPLEX: ["3", "4"],
  IMARGUMENT: ['"1+i"'],
  IMEXP: ['"1+i"'],
  IMLN: ['"1+i"'],
  IMLOG10: ['"1+i"'],
  IMLOG2: ['"1+i"'],
  IMPOWER: ['"1+i"', "2"],
  IMSQRT: ['"3+4i"'],
  IMSIN: ['"1+i"'],
  IMCOS: ['"1+i"'],
  IMTAN: ['"1+i"'],
  IMSEC: ['"1+i"'],
  IMCSC: ['"1+i"'],
  IMCOT: ['"1+i"'],
  IMSINH: ['"1+i"'],
  IMCOSH: ['"1+i"'],
  IMSECH: ['"1+i"'],
  IMCSCH: ['"1+i"'],
  BESSELI: ["1", "0"],
  BESSELJ: ["1", "0"],
  BESSELK: ["1", "0"],
  BESSELY: ["1", "0"],
  "ERFC.PRECISE": ["1"],
  "NORM.DIST": ["1", "0", "1", "TRUE"],
  "NORM.INV": ["0.5", "0", "1"],
  "T.TEST": [formulaTemplateReference, formulaTemplateReference, "2", "2"],
  "CHISQ.TEST": [formulaTemplateReference, formulaTemplateReference],
  BETADIST: ["0.5", "1", "1"],
  BETAINV: ["0.5", "1", "1"],
  BINOMDIST: ["1", "10", "0.5", "TRUE"],
  CHIDIST: ["1", "1"],
  CHIDISTRT: ["1", "1"],
  CHIINV: ["0.5", "1"],
  CHIINVRT: ["0.5", "1"],
  CHITEST: [formulaTemplateReference, formulaTemplateReference],
  CONFIDENCE: ["0.05", "1", "30"],
  COVAR: [formulaTemplateReference, formulaTemplateReference],
  COVARIANCEP: [formulaTemplateReference, formulaTemplateReference],
  COVARIANCES: [formulaTemplateReference, formulaTemplateReference],
  CRITBINOM: ["10", "0.5", "0.5"],
  EXPONDIST: ["1", "1", "TRUE"],
  FDIST: ["1", "1", "1"],
  FDISTRT: ["1", "1", "1"],
  FINV: ["0.5", "1", "1"],
  FINVRT: ["0.5", "1", "1"],
  FTEST: [formulaTemplateReference, formulaTemplateReference],
  GAMMADIST: ["1", "1", "1", "TRUE"],
  GAMMAINV: ["0.5", "1", "1"],
  HYPGEOMDIST: ["1", "10", "5", "20"],
  LOGINV: ["0.5", "0", "1"],
  LOGNORMDIST: ["1", "0", "1"],
  LOGNORMINV: ["0.5", "0", "1"],
  NEGBINOMDIST: ["1", "1", "0.5"],
  NORMDIST: ["1", "0", "1", "TRUE"],
  NORMINV: ["0.5", "0", "1"],
  NORMSDIST: ["1"],
  NORMSINV: ["0.5"],
  POISSON: ["1", "1", "TRUE"],
  POISSONDIST: ["1", "1", "TRUE"],
  STDEV: [formulaTemplateReference],
  STDEVP: [formulaTemplateReference],
  TINV: ["0.5", "1"],
  TINV2T: ["0.5", "1"],
  TDIST: ["1", "1", "2"],
  TDIST2T: ["1", "1"],
  TDISTRT: ["1", "1"],
  TTEST: [formulaTemplateReference, formulaTemplateReference, "2", "2"],
  VAR: [formulaTemplateReference],
  VARP: [formulaTemplateReference],
  WEIBULL: ["1", "1", "1", "TRUE"],
  WEIBULLDIST: ["1", "1", "1", "TRUE"],
  ZTEST: [formulaTemplateReference, () => firstCellReferenceFromRangeLabel(formulaTemplateReference())],
  DSUM: [formulaTemplateReference, "1", formulaTemplateReference],
  DCOUNT: [formulaTemplateReference, "1", formulaTemplateReference],
  DCOUNTA: [formulaTemplateReference, "1", formulaTemplateReference],
  DAVERAGE: [formulaTemplateReference, "1", formulaTemplateReference],
  DMAX: [formulaTemplateReference, "1", formulaTemplateReference],
  DMIN: [formulaTemplateReference, "1", formulaTemplateReference],
  DPRODUCT: [formulaTemplateReference, "1", formulaTemplateReference],
  DSTDEV: [formulaTemplateReference, "1", formulaTemplateReference],
  DSTDEVP: [formulaTemplateReference, "1", formulaTemplateReference],
  DVAR: [formulaTemplateReference, "1", formulaTemplateReference],
  DVARP: [formulaTemplateReference, "1", formulaTemplateReference],
};



const FORMULA_ARGUMENT_DESCRIPTIONS = {
  IF: {
    logical_test: "TRUE または FALSE に評価される条件を指定します。",
    value_if_true: "条件が TRUE の場合に返す値を指定します。",
    value_if_false: "条件が FALSE の場合に返す値を指定します。",
  },
  IFS: {
    logical_test1: "TRUE または FALSE に評価される最初の条件を指定します。",
    value_if_true1: "対応する条件が TRUE の場合に返す値を指定します。",
  },
  LET: {
    name1: "数式内で使用する名前を指定します。",
    name_value1: "名前に割り当てる値、参照、または数式を指定します。",
    calculation_or_name2: "最後の引数では計算式を指定します。追加の名前を定義する場合は名前を指定します。",
    name_value2: "追加した名前に割り当てる値、参照、または数式を指定します。",
  },
  LAMBDA: {
    parameter1: "計算式で使用するパラメーター名を指定します。",
    calculation: "パラメーターを使って評価する計算式を指定します。",
  },
  MAKEARRAY: {
    rows: "作成する配列の行数を指定します。",
    cols: "作成する配列の列数を指定します。",
    lambda: "行番号と列番号を受け取る LAMBDA を指定します。",
  },
  BYROW: {
    array: "行ごとに処理する配列または範囲を指定します。",
    lambda: "各行を受け取る LAMBDA を指定します。",
  },
  BYCOL: {
    array: "列ごとに処理する配列または範囲を指定します。",
    lambda: "各列を受け取る LAMBDA を指定します。",
  },
  MAP: {
    array1: "処理する最初の配列または範囲を指定します。",
    lambda_or_array2: "最後の引数では LAMBDA を指定します。複数配列の場合は追加の配列も指定します。",
  },
  REDUCE: {
    initial_value: "累積計算の初期値を指定します。",
    array: "累積する配列または範囲を指定します。",
    lambda: "累積値と現在値を受け取る LAMBDA を指定します。",
  },
  SCAN: {
    initial_value: "累積計算の初期値を指定します。",
    array: "累積する配列または範囲を指定します。",
    lambda: "累積値と現在値を受け取る LAMBDA を指定します。",
  },
  XOR: {
    logical1: "TRUE または FALSE に評価される論理値を指定します。",
  },
  AND: {
    logical1: "TRUE または FALSE に評価される最初の条件を指定します。",
    logical2: "追加で評価する条件を指定します。",
  },
  OR: {
    logical1: "TRUE または FALSE に評価される最初の条件を指定します。",
    logical2: "追加で評価する条件を指定します。",
  },
  NOT: {
    logical: "反転する論理値または条件を指定します。",
  },
  IFERROR: {
    value: "エラーかどうかを確認する値または数式を指定します。",
    value_if_error: "数式がエラーの場合に返す値を指定します。",
  },
  SWITCH: {
    expression: "比較する値または式を指定します。",
    value1: "一致を確認する最初の値を指定します。",
    result1: "対応する値に一致した場合に返す結果を指定します。",
    default_or_value2: "一致しない場合の既定値、または追加の比較値を指定します。",
  },
  CHOOSE: {
    index_num: "返す値を選ぶ番号を指定します。",
    value1: "1 が指定された場合に返す値を指定します。",
    value2: "2 が指定された場合に返す値を指定します。",
  },
  XLOOKUP: {
    lookup_value: "検索する値を指定します。",
    lookup_array: "検索値を探す範囲または配列を指定します。",
    return_array: "戻り値を取得する範囲または配列を指定します。",
    if_not_found: "一致しない場合に返す値を指定します。",
    match_mode: "一致の種類を指定します。",
    search_mode: "検索方向を指定します。",
  },
  VLOOKUP: {
    lookup_value: "検索する値を指定します。",
    table_array: "検索対象の表範囲を指定します。",
    col_index_num: "戻り値を含む列番号を指定します。",
    range_lookup: "近似一致なら TRUE、完全一致なら FALSE を指定します。",
  },
  HLOOKUP: {
    lookup_value: "検索する値を指定します。",
    table_array: "検索対象の表範囲を指定します。",
    row_index_num: "戻り値を含む行番号を指定します。",
    range_lookup: "近似一致なら TRUE、完全一致なら FALSE を指定します。",
  },
  INDEX: {
    array: "値を取り出す範囲または配列を指定します。",
    row_num: "取り出す行番号を指定します。",
    column_num: "取り出す列番号を指定します。",
  },
  MATCH: {
    lookup_value: "検索する値を指定します。",
    lookup_array: "検索値を探す範囲を指定します。",
    match_type: "照合の種類を指定します。0 は完全一致です。",
  },
  FILTER: {
    array: "抽出する元データの範囲または配列を指定します。",
    include: "抽出する行または列を TRUE/FALSE で示す条件配列を指定します。",
    if_empty: "該当するデータがない場合に返す値を指定します。",
  },
  SORT: {
    array: "並べ替える範囲または配列を指定します。",
    sort_index: "並べ替えの基準にする行番号または列番号を指定します。",
    sort_order: "1 は昇順、-1 は降順を指定します。",
    by_col: "列方向に並べ替える場合は TRUE を指定します。",
  },
  SORTBY: {
    array: "並べ替える範囲または配列を指定します。",
    by_array1: "並べ替えの基準にする範囲または配列を指定します。",
    sort_order1: "1 は昇順、-1 は降順を指定します。",
  },
  UNIQUE: {
    array: "一意の値を取り出す範囲または配列を指定します。",
    by_col: "列方向で比較する場合は TRUE を指定します。",
    exactly_once: "1 回だけ出現する値だけを返す場合は TRUE を指定します。",
  },
  SEQUENCE: {
    rows: "作成する配列の行数を指定します。",
    columns: "作成する配列の列数を指定します。",
    start: "連続値の開始値を指定します。",
    step: "隣り合う値の増分を指定します。",
  },
  OFFSET: {
    reference: "基準にするセルまたはセル範囲を指定します。",
    rows: "基準から移動する行数を指定します。",
    cols: "基準から移動する列数を指定します。",
    height: "返す参照の行数を指定します。",
    width: "返す参照の列数を指定します。",
  },
  INDIRECT: {
    ref_text: "参照を表す文字列を指定します。",
    a1: "A1 形式なら TRUE、R1C1 形式なら FALSE を指定します。",
  },
  ADDRESS: {
    row_num: "セル参照の行番号を指定します。",
    column_num: "セル参照の列番号を指定します。",
    abs_num: "絶対参照と相対参照の種類を指定します。",
    a1: "A1 形式なら TRUE、R1C1 形式なら FALSE を指定します。",
    sheet_text: "シート名を指定します。",
  },
  SUM: {
    number1: "合計する最初の数値またはセル範囲を指定します。",
    number2: "合計に追加する数値またはセル範囲を指定します。",
  },
  SUMPRODUCT: {
    array1: "積を求めて合計する最初の配列または範囲を指定します。",
    array2: "追加で掛け合わせる配列または範囲を指定します。",
  },
  SUMIFS: {
    sum_range: "実際に合計するセル範囲を指定します。",
    criteria_range1: "条件を評価するセル範囲を指定します。",
    criteria1: "合計対象を決める条件を指定します。",
  },
  SUMIF: {
    range: "条件を評価するセル範囲を指定します。",
    criteria: "合計対象を決める条件を指定します。",
    sum_range: "実際に合計するセル範囲を指定します。",
  },
  AGGREGATE: {
    function_num: "集計方法を 1 から 19 の番号で指定します。1 は AVERAGE、9 は SUM、14 は LARGE です。",
    options: "無視する値を 0 から 7 の番号で指定します。6 はエラー値を無視します。",
    ref1: "集計対象の最初の参照、配列、または値を指定します。",
    ref2: "追加の参照、または function_num が 14 から 19 の場合は k や quart を指定します。",
  },
  COUNTIFS: {
    criteria_range1: "条件を評価するセル範囲を指定します。",
    criteria1: "数える対象を決める条件を指定します。",
  },
  COUNTIF: {
    range: "条件を評価するセル範囲を指定します。",
    criteria: "数える対象を決める条件を指定します。",
  },
  COUNT: {
    value1: "数値を含むセル、範囲、または値を指定します。",
    value2: "追加で数えるセル、範囲、または値を指定します。",
  },
  COUNTA: {
    value1: "空白以外を数えるセル、範囲、または値を指定します。",
    value2: "追加で数えるセル、範囲、または値を指定します。",
  },
  COUNTBLANK: {
    range: "空白セルを数える範囲を指定します。",
  },
  AVERAGE: {
    number1: "平均を求める最初の数値またはセル範囲を指定します。",
    number2: "平均に含める追加の数値またはセル範囲を指定します。",
  },
  AVERAGEIF: {
    range: "条件を評価するセル範囲を指定します。",
    criteria: "平均対象を決める条件を指定します。",
    average_range: "実際に平均を求めるセル範囲を指定します。",
  },
  AVERAGEIFS: {
    average_range: "平均を求めるセル範囲を指定します。",
    criteria_range1: "条件を評価するセル範囲を指定します。",
    criteria1: "平均対象を決める条件を指定します。",
  },
  MIN: {
    number1: "最小値を求める最初の数値またはセル範囲を指定します。",
    number2: "最小値の判定に含める追加の数値またはセル範囲を指定します。",
  },
  MAX: {
    number1: "最大値を求める最初の数値またはセル範囲を指定します。",
    number2: "最大値の判定に含める追加の数値またはセル範囲を指定します。",
  },
  SUBTOTAL: {
    function_num: "集計方法を番号で指定します。9 は SUM、1 は AVERAGE です。",
    ref1: "集計対象の最初の範囲を指定します。",
    ref2: "追加で集計する範囲を指定します。",
  },
  ROUND: {
    number: "丸める数値を指定します。",
    num_digits: "丸めた後の桁数を指定します。",
  },
  ROUNDUP: {
    number: "切り上げる数値を指定します。",
    num_digits: "切り上げた後の桁数を指定します。",
  },
  ROUNDDOWN: {
    number: "切り下げる数値を指定します。",
    num_digits: "切り下げた後の桁数を指定します。",
  },
  PRODUCT: {
    number1: "積を求める最初の数値またはセル範囲を指定します。",
    number2: "積に含める追加の数値またはセル範囲を指定します。",
  },
  MOD: {
    number: "割られる数値を指定します。",
    divisor: "割る数値を指定します。",
  },
  ABS: {
    number: "絶対値を求める数値を指定します。",
  },
  INT: {
    number: "切り捨てる数値を指定します。",
  },
  CEILING: {
    number: "基準値の倍数に切り上げる数値を指定します。",
    significance: "切り上げの基準値を指定します。",
  },
  FLOOR: {
    number: "基準値の倍数に切り下げる数値を指定します。",
    significance: "切り下げの基準値を指定します。",
  },
  MROUND: {
    number: "倍数に丸める数値を指定します。",
    multiple: "丸めの基準にする倍数を指定します。",
  },
  LARGE: {
    array: "対象の数値範囲または配列を指定します。",
    k: "何番目に大きい値を返すか指定します。",
  },
  SMALL: {
    array: "対象の数値範囲または配列を指定します。",
    k: "何番目に小さい値を返すか指定します。",
  },
  RANK: {
    number: "順位を求める数値を指定します。",
    ref: "順位を比較する数値範囲を指定します。",
    order: "0 または省略で降順、0 以外で昇順を指定します。",
  },
  "RANK.EQ": {
    number: "順位を求める数値を指定します。",
    ref: "順位を比較する数値範囲を指定します。",
    order: "0 または省略で降順、0 以外で昇順を指定します。",
  },
  "RANK.AVG": {
    number: "順位を求める数値を指定します。",
    ref: "順位を比較する数値範囲を指定します。",
    order: "0 または省略で降順、0 以外で昇順を指定します。",
  },
  STDEV: {
    number1: "標本標準偏差を求める最初の数値または範囲を指定します。",
    number2: "追加の数値または範囲を指定します。",
  },
  "STDEV.S": {
    number1: "標本標準偏差を求める最初の数値または範囲を指定します。",
    number2: "追加の数値または範囲を指定します。",
  },
  "STDEV.P": {
    number1: "母集団標準偏差を求める最初の数値または範囲を指定します。",
    number2: "追加の数値または範囲を指定します。",
  },
  VAR: {
    number1: "標本分散を求める最初の数値または範囲を指定します。",
    number2: "追加の数値または範囲を指定します。",
  },
  "VAR.S": {
    number1: "標本分散を求める最初の数値または範囲を指定します。",
    number2: "追加の数値または範囲を指定します。",
  },
  "VAR.P": {
    number1: "母集団分散を求める最初の数値または範囲を指定します。",
    number2: "追加の数値または範囲を指定します。",
  },
  DATE: {
    year: "年を表す数値を指定します。",
    month: "月を表す数値を指定します。",
    day: "日を表す数値を指定します。",
  },
  TIME: {
    hour: "時を表す数値を指定します。",
    minute: "分を表す数値を指定します。",
    second: "秒を表す数値を指定します。",
  },
  EDATE: {
    start_date: "基準にする日付を指定します。",
    months: "前後に移動する月数を指定します。",
  },
  EOMONTH: {
    start_date: "基準にする日付を指定します。",
    months: "月末日を求める月数を指定します。",
  },
  DATEDIF: {
    start_date: "開始日を指定します。",
    end_date: "終了日を指定します。",
    unit: "差の単位を Y、M、D などで指定します。",
  },
  NETWORKDAYS: {
    start_date: "開始日を指定します。",
    end_date: "終了日を指定します。",
    holidays: "稼働日から除外する祝日や休日の範囲を指定します。",
  },
  "NETWORKDAYS.INTL": {
    start_date: "開始日を指定します。",
    end_date: "終了日を指定します。",
    weekend: "週末の曜日を番号または文字列で指定します。",
    holidays: "稼働日から除外する祝日や休日の範囲を指定します。",
  },
  WORKDAY: {
    start_date: "開始日を指定します。",
    days: "開始日から数える稼働日数を指定します。",
    holidays: "稼働日から除外する祝日や休日の範囲を指定します。",
  },
  "WORKDAY.INTL": {
    start_date: "開始日を指定します。",
    days: "開始日から数える稼働日数を指定します。",
    weekend: "週末の曜日を番号または文字列で指定します。",
    holidays: "稼働日から除外する祝日や休日の範囲を指定します。",
  },
  YEARFRAC: {
    start_date: "開始日を指定します。",
    end_date: "終了日を指定します。",
    basis: "日数計算の基準を指定します。",
  },
  PMT: {
    rate: "各期の利率を指定します。",
    nper: "支払回数の合計を指定します。",
    pv: "現在価値、つまり元金を指定します。",
    fv: "最後の支払い後に残る将来価値を指定します。",
    type: "支払い期日を指定します。0 は期末、1 は期首です。",
  },
  PV: {
    rate: "各期の利率を指定します。",
    nper: "支払回数の合計を指定します。",
    pmt: "各期の支払額を指定します。",
    fv: "将来価値を指定します。",
    type: "支払い期日を指定します。0 は期末、1 は期首です。",
  },
  FV: {
    rate: "各期の利率を指定します。",
    nper: "支払回数の合計を指定します。",
    pmt: "各期の支払額を指定します。",
    pv: "現在価値を指定します。",
    type: "支払い期日を指定します。0 は期末、1 は期首です。",
  },
  NPV: {
    rate: "割引率を指定します。",
    value1: "現在価値を求める最初の支払いまたは収入を指定します。",
    value2: "追加の支払いまたは収入を指定します。",
  },
  RATE: {
    nper: "支払回数の合計を指定します。",
    pmt: "各期の支払額を指定します。",
    pv: "現在価値を指定します。",
    fv: "将来価値を指定します。",
    type: "支払い期日を指定します。0 は期末、1 は期首です。",
    guess: "利率の推定値を指定します。",
  },
  NPER: {
    rate: "各期の利率を指定します。",
    pmt: "各期の支払額を指定します。",
    pv: "現在価値を指定します。",
    fv: "将来価値を指定します。",
    type: "支払い期日を指定します。0 は期末、1 は期首です。",
  },
  IPMT: {
    rate: "各期の利率を指定します。",
    per: "利息額を求める期を指定します。",
    nper: "支払回数の合計を指定します。",
    pv: "現在価値を指定します。",
    fv: "将来価値を指定します。",
    type: "支払い期日を指定します。0 は期末、1 は期首です。",
  },
  PPMT: {
    rate: "各期の利率を指定します。",
    per: "元金額を求める期を指定します。",
    nper: "支払回数の合計を指定します。",
    pv: "現在価値を指定します。",
    fv: "将来価値を指定します。",
    type: "支払い期日を指定します。0 は期末、1 は期首です。",
  },
  "LOGNORM.DIST": {
    x: "対数正規分布を評価する値を指定します。",
    mean: "ln(x) の平均を指定します。",
    standard_dev: "ln(x) の標準偏差を指定します。",
    cumulative: "累積分布関数なら TRUE、確率密度関数なら FALSE を指定します。",
  },
  "LOGNORM.INV": {
    probability: "対数正規分布に対応する確率を指定します。",
    mean: "ln(x) の平均を指定します。",
    standard_dev: "ln(x) の標準偏差を指定します。",
  },
  DELTA: {
    number1: "比較する最初の数値を指定します。",
    number2: "比較する 2 番目の数値を指定します。省略すると 0 と比較します。",
  },
  TBILLEQ: {
    settlement: "財務省短期証券の受渡日を指定します。",
    maturity: "財務省短期証券の満期日を指定します。",
    discount: "財務省短期証券の割引率を指定します。",
  },
  TEXTJOIN: {
    delimiter: "結合する文字列の間に入れる区切り文字を指定します。",
    ignore_empty: "空白セルを無視するかどうかを指定します。",
    text1: "結合する最初の文字列または範囲を指定します。",
    text2: "結合する追加の文字列または範囲を指定します。",
  },
  TEXTSPLIT: {
    text: "分割する文字列またはセル参照を指定します。",
    col_delimiter: "列方向に分割する区切り文字を指定します。",
    row_delimiter: "行方向に分割する区切り文字を指定します。",
    ignore_empty: "連続する区切り文字で空白を作らない場合は TRUE を指定します。",
    match_mode: "0 は大文字小文字を区別し、1 は区別しません。",
    pad_with: "行ごとの列数が違う場合に埋める値を指定します。",
  },
  TEXTBEFORE: {
    text: "検索対象の文字列またはセル参照を指定します。",
    delimiter: "基準にする区切り文字を指定します。",
    instance_num: "何番目の区切り文字を使うか指定します。負の数は末尾から数えます。",
    match_mode: "0 は大文字小文字を区別し、1 は区別しません。",
    match_end: "文字列の末尾も区切り文字として扱う場合は TRUE を指定します。",
    if_not_found: "区切り文字が見つからない場合に返す値を指定します。",
  },
  TEXTAFTER: {
    text: "検索対象の文字列またはセル参照を指定します。",
    delimiter: "基準にする区切り文字を指定します。",
    instance_num: "何番目の区切り文字を使うか指定します。負の数は末尾から数えます。",
    match_mode: "0 は大文字小文字を区別し、1 は区別しません。",
    match_end: "文字列の末尾も区切り文字として扱う場合は TRUE を指定します。",
    if_not_found: "区切り文字が見つからない場合に返す値を指定します。",
  },
  HYPERLINK: {
    link_location: "リンク先のアドレスまたはセル参照を指定します。",
    friendly_name: "セルに表示する文字列を指定します。",
  },
  CONCAT: {
    text1: "結合する最初の文字列または範囲を指定します。",
    text2: "結合する追加の文字列または範囲を指定します。",
  },
  CONCATENATE: {
    text1: "結合する最初の文字列を指定します。",
    text2: "結合する追加の文字列を指定します。",
  },
  LEFT: {
    text: "取り出す文字列またはセル参照を指定します。",
    num_chars: "左端から取り出す文字数を指定します。",
  },
  RIGHT: {
    text: "取り出す文字列またはセル参照を指定します。",
    num_chars: "右端から取り出す文字数を指定します。",
  },
  MID: {
    text: "取り出す文字列またはセル参照を指定します。",
    start_num: "取り出しを開始する文字位置を指定します。",
    num_chars: "取り出す文字数を指定します。",
  },
  FIND: {
    find_text: "検索する文字列を指定します。",
    within_text: "検索対象の文字列またはセル参照を指定します。",
    start_num: "検索を開始する文字位置を指定します。",
  },
  SEARCH: {
    find_text: "検索する文字列を指定します。",
    within_text: "検索対象の文字列またはセル参照を指定します。",
    start_num: "検索を開始する文字位置を指定します。",
  },
  SUBSTITUTE: {
    text: "置換対象の文字列またはセル参照を指定します。",
    old_text: "置換前の文字列を指定します。",
    new_text: "置換後の文字列を指定します。",
    instance_num: "何番目の一致だけを置換するか指定します。",
  },
  REPLACE: {
    old_text: "置換対象の文字列またはセル参照を指定します。",
    start_num: "置換を開始する文字位置を指定します。",
    num_chars: "置換する文字数を指定します。",
    new_text: "置換後の文字列を指定します。",
  },
  TEXT: {
    value: "表示形式を適用する値を指定します。",
    format_text: "適用する表示形式を文字列で指定します。",
  },
  VALUE: {
    text: "数値に変換する文字列を指定します。",
  },
  NUMBERVALUE: {
    text: "数値に変換する文字列を指定します。",
    decimal_separator: "小数点として扱う文字を指定します。",
    group_separator: "桁区切りとして扱う文字を指定します。",
  },
};



const FORMULA_REFERENCE_SHEET_TOKEN_SOURCE = "(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_. ]*)!";


const FORMULA_REFERENCE_BODY_TOKEN_SOURCE = "\\$?[A-Za-z]{1,3}\\$?[1-9][0-9]{0,6}(?::\\$?[A-Za-z]{1,3}\\$?[1-9][0-9]{0,6})?|\\$?[A-Za-z]{1,3}:\\$?[A-Za-z]{1,3}|\\$?[1-9][0-9]{0,6}:\\$?[1-9][0-9]{0,6}";



const FORMULA_ERROR_CODES = ["#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#N/A", "#NUM!", "#NULL!", "#SPILL!", "#CALC!", "#CYCLE!"];



function initWorkbookNameDisplay() {
  if (!$workbookNameDisplay.length || !$workbookNameDisplayText.length) return;
  if (!workbookNameMutationObserver && $workbookName[0] && typeof MutationObserver !== "undefined") {
    workbookNameMutationObserver = new MutationObserver(requestWorkbookNameDisplayUpdate);
    workbookNameMutationObserver.observe($workbookName[0], { childList: true, characterData: true, subtree: true });
  }
  if (!workbookNameResizeObserver && $workbookNameDisplay[0] && typeof ResizeObserver !== "undefined") {
    workbookNameResizeObserver = new ResizeObserver(requestWorkbookNameDisplayUpdate);
    const row = $workbookNameDisplay[0].closest?.(".ribbon-tab-row");
    [
      $workbookNameDisplay[0],
      row,
      row?.querySelector?.(".ribbon-tabs"),
      row?.querySelector?.(".ribbon-status"),
      row?.querySelector?.(".ribbon-display-options"),
    ]
      .filter(Boolean)
      .forEach((element) => workbookNameResizeObserver.observe(element));
  } else {
    window.addEventListener("resize", requestWorkbookNameDisplayUpdate);
  }
  updateWorkbookNameDisplay();
}



function requestWorkbookNameDisplayUpdate() {
  if (workbookNameDisplayRaf) return;
  const run = () => {
    workbookNameDisplayRaf = 0;
    updateWorkbookNameDisplay();
  };
  if (typeof window.requestAnimationFrame === "function") {
    workbookNameDisplayRaf = window.requestAnimationFrame(run);
  } else {
    workbookNameDisplayRaf = window.setTimeout(run, 16);
  }
}



function measureWorkbookNameText(text, font) {
  if (!workbookNameMeasureContext) {
    const canvas = document.createElement("canvas");
    workbookNameMeasureContext = canvas.getContext("2d");
  }
  if (!workbookNameMeasureContext) return String(text || "").length * 8;
  workbookNameMeasureContext.font = font || "700 10px sans-serif";
  return workbookNameMeasureContext.measureText(String(text || "")).width;
}



const SORT_CUSTOM_LISTS = [
  {
    id: "month-ja",
    label: "ユーザー設定リスト: 1月, 2月, 3月...",
    values: ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"],
  },
  {
    id: "weekday-ja",
    label: "ユーザー設定リスト: 日曜日, 月曜日, 火曜日...",
    values: ["日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"],
    aliases: [["日", "日曜"], ["月", "月曜"], ["火", "火曜"], ["水", "水曜"], ["木", "木曜"], ["金", "金曜"], ["土", "土曜"]],
  },
  {
    id: "month-en",
    label: "ユーザー設定リスト: January, February, March...",
    values: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
    aliases: [["Jan"], ["Feb"], ["Mar"], ["Apr"], ["May"], ["Jun"], ["Jul"], ["Aug"], ["Sep", "Sept"], ["Oct"], ["Nov"], ["Dec"]],
  },
  {
    id: "weekday-en",
    label: "ユーザー設定リスト: Sunday, Monday, Tuesday...",
    values: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
    aliases: [["Sun"], ["Mon"], ["Tue", "Tues"], ["Wed"], ["Thu", "Thur", "Thurs"], ["Fri"], ["Sat"]],
  },
];



const AUTO_FILTER_MONTH_NAME_INDEX = {
  jan: 1,
  january: 1,
  feb: 2,
  february: 2,
  mar: 3,
  march: 3,
  apr: 4,
  april: 4,
  may: 5,
  jun: 6,
  june: 6,
  jul: 7,
  july: 7,
  aug: 8,
  august: 8,
  sep: 9,
  sept: 9,
  september: 9,
  oct: 10,
  october: 10,
  nov: 11,
  november: 11,
  dec: 12,
  december: 12,
};



const FLASH_FILL_REGEX_SEGMENTS = [
  { id: "digits", regex: /\d+/g },
  { id: "letters", regex: /[A-Za-z]+/g },
  { id: "uppercase", regex: /[A-Z]+/g },
];



const DATA_SUMMARY_FUNCTIONS = [
  { value: "sum", label: "合計", subtotalCode: 9 },
  { value: "count", label: "データの個数", subtotalCode: 2 },
  { value: "average", label: "平均", subtotalCode: 1 },
  { value: "max", label: "最大値", subtotalCode: 4 },
  { value: "min", label: "最小値", subtotalCode: 5 },
  { value: "product", label: "積", subtotalCode: 6 },
  { value: "countNums", label: "数値の個数", subtotalCode: 3 },
  { value: "stdDev", label: "標準偏差", subtotalCode: 7 },
  { value: "stdDevP", label: "標準偏差P", subtotalCode: 8 },
  { value: "var", label: "分散", subtotalCode: 10 },
  { value: "varP", label: "分散P", subtotalCode: 11 },
];



const DATA_ANALYSIS_METHODS = [
  { value: "anova1", label: "分散分析: 一元配置" },
  { value: "anova2Replication", label: "分散分析: 繰り返しのある二元配置" },
  { value: "anova2NoReplication", label: "分散分析: 繰り返しのない二元配置" },
  { value: "correlation", label: "相関" },
  { value: "covariance", label: "共分散" },
  { value: "descriptive", label: "基本統計量" },
  { value: "exponentialSmoothing", label: "指数平滑" },
  { value: "fTest", label: "F 検定: 2 標本を使った分散の検定" },
  { value: "fourier", label: "フーリエ解析" },
  { value: "histogram", label: "ヒストグラム" },
  { value: "movingAverage", label: "移動平均" },
  { value: "random", label: "乱数発生" },
  { value: "rankPercentile", label: "順位と百分位数" },
  { value: "regression", label: "回帰分析" },
  { value: "sampling", label: "サンプリング" },
  { value: "tTestPaired", label: "t 検定: 一対の標本による平均の検定" },
  { value: "tTestEqual", label: "t 検定: 等分散を仮定した 2 標本による検定" },
  { value: "tTestUnequal", label: "t 検定: 分散が等しくないと仮定した 2 標本による検定" },
  { value: "zTest", label: "z 検定: 2 標本による平均の検定" },
];



const DATA_ANALYSIS_TWO_VARIABLE_METHODS = new Set(["fTest", "tTestPaired", "tTestEqual", "tTestUnequal", "zTest"]);



const DATA_ANALYSIS_HELP = {
  descriptive: "選択したデータ範囲について、平均、標準偏差、最大値、最小値、順位などの基本統計量を作成します。",
  correlation: "複数の数値系列の相関係数を計算します。入力範囲には 2 つ以上の数値系列が必要です。",
  covariance: "複数の数値系列の共分散を計算します。入力範囲には 2 つ以上の数値系列が必要です。",
  histogram: "入力範囲の値をビン範囲ごとに集計し、度数分布表を作成します。必要に応じてグラフも作成できます。",
  movingAverage: "指定した期間で移動平均を計算し、時系列データの傾向を確認します。",
  exponentialSmoothing: "指定した平滑化定数を使って指数平滑を計算し、時系列データをなめらかにします。",
  regression: "入力 Y 範囲と入力 X 範囲から回帰分析を行い、係数、決定係数、残差などを出力します。",
  anova1: "複数のグループの平均に有意な差があるかを、一元配置の分散分析で確認します。",
  anova2NoReplication: "繰り返しのない二元配置の分散分析を行い、行要因と列要因の影響を確認します。",
  anova2Replication: "繰り返しのある二元配置の分散分析を行い、要因と交互作用を確認します。",
  rankPercentile: "入力範囲の値について順位と百分位数を出力します。",
  sampling: "入力範囲から周期またはランダム方式で標本を抽出します。",
  random: "指定した分布、行数、列数に従って乱数表を作成します。",
  fTest: "2 つの標本の分散に差があるかを F 検定で確認します。",
  tTestPaired: "対応のある 2 つの標本について、平均の差を t 検定で確認します。",
  tTestEqual: "等分散を仮定した 2 標本の t 検定を行います。",
  tTestUnequal: "分散が等しくないと仮定した 2 標本の t 検定を行います。",
  zTest: "既知の分散を使って 2 標本の平均差を z 検定で確認します。",
  fourier: "入力範囲の値に対してフーリエ解析を行います。逆変換を選択することもできます。",
};



const HTML_NAMED_COLORS = {
  aqua: "#00FFFF",
  black: "#000000",
  blue: "#0000FF",
  cyan: "#00FFFF",
  fuchsia: "#FF00FF",
  gray: "#808080",
  grey: "#808080",
  green: "#008000",
  lime: "#00FF00",
  magenta: "#FF00FF",
  maroon: "#800000",
  navy: "#000080",
  olive: "#808000",
  orange: "#FFA500",
  purple: "#800080",
  red: "#FF0000",
  silver: "#C0C0C0",
  teal: "#008080",
  white: "#FFFFFF",
  window: "#FFFFFF",
  windowtext: "#000000",
  yellow: "#FFFF00",
};



const CELL_EDIT_MIN_PADDING_TOP = 2;


const CELL_EDIT_PADDING_BOTTOM = 2;



const AUTO_FILL_OPTIONS_BUTTON_SIZE = 20;


const AUTO_FILL_OPTIONS_VIEWPORT_MARGIN = 4;


const AUTO_FILL_OPTIONS_MENU_WIDTH = 190;



const AUTO_FILL_LISTS = [
  ["日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"],
  ["日", "月", "火", "水", "木", "金", "土"],
  ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],
  ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
  ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"],
  ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
  ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
  ["Q1", "Q2", "Q3", "Q4"],
  ["1Q", "2Q", "3Q", "4Q"],
  ["第1四半期", "第2四半期", "第3四半期", "第4四半期"],
  ["第1季度", "第2季度", "第3季度", "第4季度"],
];



const SAVE_AS_FILE_TYPES = [
  {
    value: "html",
    label: "WebSheet HTML (*.html)",
    extension: "html",
    mime: "text/html;charset=utf-8",
    description: "WebSheet HTML",
    pickerAccept: { "text/html": [".html"] },
  },
  {
    value: "xlsx",
    label: "Excel ブック (*.xlsx)",
    extension: "xlsx",
    mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    description: "Excel ブック",
    pickerAccept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] },
  },
  {
    value: "xlsm",
    label: "Excel マクロ有効ブック (*.xlsm)",
    extension: "xlsm",
    mime: "application/vnd.ms-excel.sheet.macroEnabled.12",
    description: "Excel マクロ有効ブック",
    pickerAccept: { "application/vnd.ms-excel.sheet.macroEnabled.12": [".xlsm"] },
  },
  {
    value: "xls",
    label: "Excel 97-2003 ブック (*.xls)",
    extension: "xls",
    mime: "application/vnd.ms-excel",
    description: "Excel 97-2003 ブック",
    pickerAccept: { "application/vnd.ms-excel": [".xls"] },
  },
];


const SAVE_AS_LOCATIONS = [
  { value: "downloads", label: "ダウンロード", detail: "ブラウザーの既定の保存先", path: "PC > ダウンロード" },
  { value: "desktop", label: "デスクトップ", detail: "保存先を参照して指定", path: "PC > デスクトップ" },
  { value: "documents", label: "ドキュメント", detail: "保存先を参照して指定", path: "PC > ドキュメント" },
  { value: "this-pc", label: "この PC", detail: "参照して保存先を選択", path: "PC" },
  { value: "websheet", label: "WebSheet", detail: "ローカル共通ランタイム用", path: "PC > WebSheet" },
];



init();

export {
  $,
  $addSheetButton,
  $appAlert,
  $bootState,
  $dropZone,
  $emptyState,
  $exportHtmlButton,
  $fileInput,
  $fileSideMenu,
  $formulaBar,
  $formulaExpandButton,
  $formulaInput,
  $formulaResizeHandle,
  $horizontalScrollSpacer,
  $horizontalScrollbar,
  $pageBackgroundInput,
  $ribbonDisplayButton,
  $ribbonDisplayMenu,
  $ribbonPanels,
  $ribbonTabs,
  $selectedCellLabel,
  $selectionKeyboardSink,
  $selectionStatus,
  $sheetFooter,
  $sheetFooterSplitter,
  $sheetHost,
  $sheetTabScrollLeft,
  $sheetTabScrollRight,
  $sheetTabs,
  $sheetWorkspace,
  $statusText,
  $toggleFormulasButton,
  $warnings,
  $workbookName,
  $workbookNameDisplay,
  $workbookNameDisplayText,
  $workbookWindowBar,
  $zoomInButton,
  $zoomLabel,
  $zoomOutButton,
  $zoomSlider,
  ANNOTATION_POINTER_HALF_HEIGHT,
  ANNOTATION_POINTER_TIP_OFFSET,
  APP_COMPANY_NAME,
  APP_DISPLAY_NAME,
  APP_ICON_DATA_URL,
  APP_ICON_VERSION,
  APP_SHORT_NAME,
  APP_TITLE,
  APP_VERSION,
  AUTO_FILL_DATE_MODES,
  AUTO_FILL_LISTS,
  AUTO_FILL_MODES,
  AUTO_FILL_OPTIONS_BUTTON_SIZE,
  AUTO_FILL_OPTIONS_MENU_WIDTH,
  AUTO_FILL_OPTIONS_VIEWPORT_MARGIN,
  AUTO_FILL_TREND_MODES,
  AUTO_FILTER_HEADER_SCAN_LIMIT,
  AUTO_FILTER_MONTH_NAME_INDEX,
  AUTO_FILTER_VALUE_LIMIT,
  AUTO_STANDALONE_RUNTIME_BASE_URL,
  AlignCenter,
  AlignHorizontalJustifyCenter,
  AlignHorizontalJustifyEnd,
  AlignHorizontalJustifyStart,
  AlignJustify,
  AlignLeft,
  AlignRight,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  ArraySize,
  ArrowDownNarrowWide,
  ArrowDownToLine,
  ArrowUpDown,
  BadgePlus,
  BarChart3,
  BookOpen,
  Box,
  CELL_BORDER_SIDES,
  CELL_EDIT_MIN_PADDING_TOP,
  CELL_EDIT_PADDING_BOTTOM,
  CELL_EDIT_TRAILING_LINE_SENTINEL,
  CHART_COLORS,
  CJK_FONT_FALLBACKS,
  COLUMN_HEADER_HEIGHT,
  CUSTOM_SHAPE_TYPE,
  CWS_AGENT_BORDER_PRESETS,
  CWS_AGENT_CONTRACT_RANGE_OPS,
  CWS_AGENT_MAX_ADVANCED_FILTER_CELLS,
  CWS_AGENT_MAX_CHART_CELLS,
  CWS_AGENT_MAX_CHART_SIZE,
  CWS_AGENT_MAX_COLUMN_WIDTH,
  CWS_AGENT_MAX_COMMENT_TEXT,
  CWS_AGENT_MAX_CONSOLIDATE_CELLS,
  CWS_AGENT_MAX_COPY_CELLS,
  CWS_AGENT_MAX_DATA_TABLE_CELLS,
  CWS_AGENT_MAX_DEDUP_CELLS,
  CWS_AGENT_MAX_FILTER_CELLS,
  CWS_AGENT_MAX_FILTER_VALUES,
  CWS_AGENT_MAX_FILTER_VALUE_TEXT,
  CWS_AGENT_MAX_FIND_REPLACE_CELLS,
  CWS_AGENT_MAX_FORECAST_CELLS,
  CWS_AGENT_MAX_FORMAT_CELLS,
  CWS_AGENT_MAX_FREEZE_COLUMNS,
  CWS_AGENT_MAX_FREEZE_ROWS,
  CWS_AGENT_MAX_GOAL_SEEK_ABS_TARGET,
  CWS_AGENT_MAX_HYPERLINK_TEXT,
  CWS_AGENT_MAX_NOTE_TEXT,
  CWS_AGENT_MAX_OPS,
  CWS_AGENT_MAX_PIVOT_CELLS,
  CWS_AGENT_MAX_QUICK_ANALYSIS_CELLS,
  CWS_AGENT_MAX_RANGE_CELLS,
  CWS_AGENT_MAX_ROW_HEIGHT,
  CWS_AGENT_MAX_SHAPE_SIZE,
  CWS_AGENT_MAX_SHAPE_TEXT,
  CWS_AGENT_MAX_SORT_CELLS,
  CWS_AGENT_MAX_SPARKLINE_SOURCE_CELLS,
  CWS_AGENT_MAX_STRUCTURE_COLUMNS,
  CWS_AGENT_MAX_STRUCTURE_ROWS,
  CWS_AGENT_MAX_TABLE_CELLS,
  CWS_AGENT_MAX_TEXT_TO_COLUMNS_CELLS,
  CWS_AGENT_MAX_VALIDATION_CELLS,
  CWS_AGENT_MAX_VALIDATION_OPTIONS,
  CWS_AGENT_MAX_VALIDATION_OPTION_TEXT,
  CWS_AGENT_MAX_WORKBOOK_TEMPLATE_CELLS,
  CWS_AGENT_MAX_WORKBOOK_TEMPLATE_SHEETS,
  CWS_AGENT_MIN_CHART_SIZE,
  CWS_AGENT_MIN_COLUMN_WIDTH,
  CWS_AGENT_MIN_ROW_HEIGHT,
  CWS_AGENT_MIN_SHAPE_SIZE,
  CWS_AGENT_PROGRESSIVE_PREVIEW_DELAY_MS,
  CWS_AI_GUIDE_URL,
  CWS_CLIPBOARD_DATA_MIME,
  CWS_CLIPBOARD_DATA_WEB_MIME,
  CWS_CLIPBOARD_HTML_MARKER,
  CWS_CLIPBOARD_RANGE_MARKER,
  CWS_CLIPBOARD_RANGE_MIME,
  CWS_HTML_FORMAT,
  CWS_MODEL_SCHEMA_URL,
  CWS_SCHEMA_VERSION,
  CWS_WEB_APP_URL,
  CalendarClock,
  CanvasRenderer,
  CaseSensitive,
  CellError,
  ChartColumn,
  ChartColumnBig,
  ChartLine,
  ChartNoAxesColumn,
  ChartPie,
  Check,
  ChevronDown,
  ChevronUp,
  ChevronsDown,
  ChevronsUp,
  CircleHelp,
  ClipboardPaste,
  Clock3,
  Columns2,
  Columns3,
  Copy,
  Crop,
  DATA_ANALYSIS_HELP,
  DATA_ANALYSIS_METHODS,
  DATA_ANALYSIS_TWO_VARIABLE_METHODS,
  DATA_OPERATION_CELL_LIMIT,
  DATA_SUMMARY_FUNCTIONS,
  DEFAULT_CELL_CSS_KEYS,
  DEFAULT_MACRO_SECURITY_LEVEL,
  DEFAULT_RECENT_SHAPE_TYPES,
  DEFAULT_SHAPE_OUTLINE_WIDTH,
  DEFAULT_SHAPE_TYPE,
  DEFAULT_STANDALONE_REMOTE_RUNTIME_BASE_URL,
  DEFAULT_STANDALONE_RUNTIME_BASE_URL,
  DEFAULT_STATUS_BAR_ITEMS,
  DEFAULT_THEME_COLORS,
  DIALOG_CHROME_ROOT_SELECTOR,
  DIALOG_CHROME_TITLE_SELECTOR,
  DRAW_COLOR_PRESETS,
  DRAW_INK_ERASERS,
  DRAW_INK_TOOLS,
  Database,
  Delete,
  Download,
  EMPTY_CELL_BORDERS,
  EMPTY_CELL_CSS,
  EMU_PER_PIXEL,
  EXCELJS_IMPORT_IGNORE_NODES,
  EXCEL_DEFAULT_PAGE_MARGINS,
  EXCEL_FONT_FALLBACKS,
  EXCEL_MAX_COLS,
  EXCEL_MAX_ROWS,
  Eraser,
  ErrorType,
  ExcelJS,
  ExcelJSModule,
  Eye,
  EyeOff,
  FALLBACK_VISIBLE_COLS,
  FALLBACK_VISIBLE_ROWS,
  FLASH_FILL_REGEX_SEGMENTS,
  FORMAT_DIALOG_DATE_FORMATS,
  FORMULA_ARGUMENT_DEFAULT_OVERRIDES,
  FORMULA_ARGUMENT_DESCRIPTIONS,
  FORMULA_BAR_DEFAULT_EXPANDED_LINES,
  FORMULA_BAR_MAX_LINES,
  FORMULA_BAR_MIN_LINES,
  FORMULA_COMPATIBILITY_FUNCTIONS,
  FORMULA_ENGINE_ARGUMENT_OVERRIDES,
  FORMULA_ENGINE_DESCRIPTION_OVERRIDES,
  FORMULA_ENGINE_STATIC_CELL_INFO_TYPES,
  FORMULA_ENGINE_SYNTAX_OVERRIDES,
  FORMULA_ERROR_CODES,
  FORMULA_REFERENCE_BODY_TOKEN_SOURCE,
  FORMULA_REFERENCE_COLORS,
  FORMULA_REFERENCE_SHEET_TOKEN_SOURCE,
  FORMULA_RIBBON_FUNCTION_MENUS,
  FilePlus,
  FileSpreadsheet,
  FileText,
  FileUp,
  Filter,
  FolderOpen,
  FunctionArgumentType,
  FunctionPlugin,
  FunctionSquare,
  GitBranch,
  Globe,
  Grid2X2,
  HISTORY_LIMIT,
  HTML_NAMED_COLORS,
  HYPERFORMULA_LICENSE,
  Hash,
  Highlighter,
  HyperFormula,
  IMAGE_RENDER_BUFFER_PX,
  IMPORTED_RANGE_DIMENSION_EXPANSION_CELL_LIMIT,
  INDEXED_COLORS,
  INFERRED_DEFAULT_CELL_CSS_KEYS,
  INSERT_FUNCTION_CATEGORY_ORDER,
  INSERT_FUNCTION_RECENT_NAMES,
  INSERT_SYMBOL_CHOICES,
  INTERNAL_CLIPBOARD_STORAGE_KEY,
  ImageIcon,
  Info,
  Italic,
  JSZip,
  JSZipModule,
  KeyRound,
  LEGACY_XLS_BORDER_STYLE_MAP,
  LEGACY_XLS_BUILTIN_NUM_FMTS,
  LEGACY_XLS_CELL_RECORDS_WITH_XF,
  LEGACY_XLS_FILL_PATTERN_NONE,
  LINE_CONNECTION_RETENTION_TOLERANCE_PX,
  LINE_SHAPE_ARROW_PADDING_EXTRA_PX,
  LINE_SHAPE_BASE_PADDING_PX,
  LINE_SHAPE_MIN_BOX_SIZE,
  LINE_SHAPE_TYPES,
  LINK_DIALOG_KINDS,
  LOCALIZED_DEFAULT_CELL_CSS_KEYS,
  Lasso,
  LinkIcon,
  ListChecks,
  ListFilter,
  LogOut,
  MACRO_SECURITY_LEVELS,
  MACRO_SECURITY_NAV_ITEMS,
  MACRO_SECURITY_OPTIONS,
  MAX_FUNCTION_ARGUMENT_ROWS,
  MAX_RECENT_FORMULA_FUNCTIONS,
  MERGE_CONTENT_WARNING_MESSAGE,
  MailCheck,
  MapIcon,
  Maximize,
  MessageSquareText,
  MessageSquareX,
  Minus,
  MonitorUp,
  MousePointer2,
  MoveRight,
  OFFLINE_TRANSLATION_DICTIONARIES,
  Omega,
  PAGE_LAYOUT_EFFECT_PRESETS,
  PAGE_LAYOUT_MARGIN_PRESETS,
  PAGE_LAYOUT_PAPER_SIZES,
  PAGE_LAYOUT_THEMES,
  PENDING_FILE_LOAD_RECOVERY_STORAGE_KEY,
  PERSISTENCE_CURRENT_KEY,
  PERSISTENCE_DB_NAME,
  PERSISTENCE_DB_VERSION,
  PERSISTENCE_IDLE_TIMEOUT_MS,
  PERSISTENCE_SAVE_DEBOUNCE_MS,
  PERSISTENCE_SNAPSHOT_VERSION,
  PERSISTENCE_STORE_NAME,
  PaintBucket,
  Paintbrush,
  Palette,
  PanelTop,
  PenLine,
  Pencil,
  PlaySquare,
  Plug,
  Printer,
  Puzzle,
  QUICK_ANALYSIS_ACTIONS,
  QUICK_ANALYSIS_TABS,
  RECENT_DATA_SOURCES_LIMIT,
  RECENT_DATA_SOURCES_STORAGE_KEY,
  RECENT_DATA_SOURCE_TEXT_LIMIT,
  RIBBON_BORDER_PRESETS,
  RIBBON_COLOR_STYLE,
  RIBBON_COLOR_TINTS,
  RIBBON_LAYOUTS,
  RIBBON_LAYOUT_STORAGE_KEY,
  RIBBON_PANEL_FOCUSABLE_SELECTOR,
  RIBBON_SHAPE_GROUPS,
  RIBBON_SHAPE_PRESET_MAP,
  RIBBON_STANDARD_COLORS,
  RIBBON_UNDERLINE_PRESETS,
  RIBBON_VISIBILITY_MODES,
  RIBBON_VISIBILITY_STORAGE_KEY,
  ROW_HEADER_DIGIT_WIDTH,
  ROW_HEADER_MAX_WIDTH,
  ROW_HEADER_MIN_WIDTH,
  ROW_HEADER_PADDING_WIDTH,
  RefreshCw,
  Repeat,
  RotateCcw,
  RotateCw,
  Ruler,
  SAVE_AS_FILE_TYPES,
  SAVE_AS_LOCATIONS,
  SCROLL_EXTEND_COLS,
  SCROLL_EXTEND_ROWS,
  SCROLL_EXTEND_THRESHOLD,
  SCROLL_TRIM_BUFFER_PX,
  SCROLL_TRIM_IDLE_MS,
  SELECTION_AUTOSCROLL_EDGE_PX,
  SELECTION_AUTOSCROLL_MAX_STEP_PX,
  SELECTION_AUTOSCROLL_MIN_STEP_PX,
  SERVER_DOWNLOAD_PAGE_PATH,
  SERVER_VERSION_INFO_PATH,
  SHAPE_ARROW_PRESETS,
  SHAPE_ARROW_SIZE_PRESETS,
  SHAPE_ARROW_TYPE_PRESETS,
  SHAPE_FILL_PATTERN_PRESETS,
  SHAPE_FIT_INSET,
  SHAPE_GLOW_PRESETS,
  SHAPE_LINE_CAP_PRESETS,
  SHAPE_LINE_COMPOUND_PRESETS,
  SHAPE_LINE_DASH_PRESETS,
  SHAPE_LINE_JOIN_PRESETS,
  SHAPE_OUTLINE_WEIGHTS,
  SHAPE_REFLECTION_PRESETS,
  SHAPE_SOFT_EDGE_PRESETS,
  SHAPE_STYLE_PRESETS,
  SHAPE_TEXT_EDITOR_Z_INDEX,
  SHEET_INSERT_TEMPLATES,
  SHEET_TAB_COLORS,
  SORT_CUSTOM_LISTS,
  SPECIAL_SHEET_INSERT_META,
  STANDALONE_IMAGE_ASSET_ENCODING_GZIP_BASE64,
  STANDALONE_IMAGE_ASSET_GZIP_MIN_BASE64_LENGTH,
  STANDALONE_IMAGE_ASSET_GZIP_MIN_SAVINGS_RATIO,
  STANDALONE_IMAGE_DEVICE_PIXEL_RATIO,
  STANDALONE_IMAGE_MAX_DIMENSION,
  STANDALONE_IMAGE_WEBP_QUALITY,
  STANDALONE_REMOTE_RUNTIME_BASE_URL_STORAGE_KEY,
  STANDALONE_RUNTIME_BASE_URL_STORAGE_KEY,
  STANDALONE_RUNTIME_CHANNEL,
  STANDALONE_RUNTIME_VERSION,
  STATUS_BAR_ITEMS,
  STATUS_BAR_STORAGE_KEY,
  STYLE_INTERN_POOL_LIMIT,
  Save,
  SaveAll,
  ScatterChart,
  Scissors,
  Search,
  Send,
  Settings,
  Shapes,
  Sigma,
  Signature,
  SimpleRangeValue,
  Sparkles,
  Square,
  SquareDashed,
  SquareFunction,
  TEXT_OVERFLOW_LOOKBACK_MAX_PX,
  TEXT_OVERFLOW_LOOKBACK_MIN_PX,
  TEXT_OVERFLOW_LOOKBACK_VIEWPORT_FACTOR,
  TRANSLATION_CELL_LIMIT,
  TRANSLATION_LANGUAGES,
  TRANSLATION_PREVIEW_LIMIT,
  TRANSLATION_SCAN_LIMIT,
  Table2,
  TableCellsMerge,
  TableCellsSplit,
  TableProperties,
  TextCursorInput,
  Trash2,
  Type,
  Underline,
  Unlink,
  UnlockKeyhole,
  Upload,
  UploadCloud,
  VERSION_INFO_FALLBACK_VERSION,
  VIRTUAL_COL_OVERSCAN,
  VIRTUAL_ROW_OVERSCAN,
  WEBSHEET_CONVERT_UNITS,
  WEBSHEET_FORMULA_ANY,
  WEBSHEET_FORMULA_BOOLEAN,
  WEBSHEET_FORMULA_ENGINE_CACHEABLE_SPILL_NAMES,
  WEBSHEET_FORMULA_INTEGER,
  WEBSHEET_FORMULA_NA_SENTINEL,
  WEBSHEET_FORMULA_NUMBER,
  WEBSHEET_FORMULA_RANGE,
  WEBSHEET_FORMULA_SCALAR,
  WEBSHEET_FORMULA_STRING,
  WEBSHEET_FORMULA_TRANSLATIONS,
  WEBSHEET_ROMAN_VALUES,
  WORKBOOK_MODEL_VERSION,
  WORKBOOK_OPEN_TRANSFER_DB_NAME,
  WORKBOOK_OPEN_TRANSFER_DB_VERSION,
  WORKBOOK_OPEN_TRANSFER_MAX_AGE_MS,
  WORKBOOK_OPEN_TRANSFER_QUERY_PARAM,
  WORKBOOK_OPEN_TRANSFER_STORE_NAME,
  WORKBOOK_OPTIONAL_RIBBON_COMMANDS,
  WORKSHEET_IMAGE_OBJECT_URL_MIN_DATA_URL_LENGTH,
  WORKSHEET_OBJECT_LAYER_BASE_Z,
  WORKSHEET_OBJECT_LAYER_STEP_Z,
  WORKSHEET_OBJECT_SELECTION_Z_OFFSET,
  WebSheetFormulaPlugin,
  WrapText,
  X,
  XLSX,
  XLSXModule,
  ZOOM_DIALOG_PRESETS,
  ZoomIn,
  cellStyleInternPools,
  createIcons,
  cwsAgentOperationSources,
  excelJsModulePromise,
  formulaEngineAutocompleteFunctionCache,
  formulaEngineAutocompleteFunctionDefinitions,
  formulaSpillAnchorCacheBySheet,
  icons,
  initCwsAgentBridge,
  initWorkbookNameDisplay,
  jsZipModulePromise,
  loadExcelJsModule,
  loadJsZipModule,
  loadXlsxModule,
  measureWorkbookNameText,
  requestWorkbookNameDisplayUpdate,
  sheetBorderCellCacheBySheet,
  sheetMergeMapsCacheByMerges,
  sheetMinimumExtentCacheBySheet,
  sheetPersistentNoteCellCacheBySheet,
  sheetStyleRangeCacheBySheet,
  sheetUsedRangeCacheBySheet,
  state,
  workbookNameDisplayRaf,
  workbookNameMeasureContext,
  workbookNameMutationObserver,
  workbookNameResizeObserver,
  worksheetImageRenderObjectUrls,
  xlsxModulePromise,
};
