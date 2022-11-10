import pandas as pd
import sqlite3

# アップロードに用いるdbmap
excel_db_map = {
    'table_name': 'customer_infomation',
    'column': {
        'No': {'type': 'text', 'index': 1},
        'お名前': {'type': 'text', 'index': 2},
        '日付_1日目': {'type': 'text', 'index': 3},
        '日付_2日目': {'type': 'text', 'index': 4},
        '日付_3日目': {'type': 'text', 'index': 5},
        '日付_4日目': {'type': 'text', 'index': 6},
        '日付_5日目': {'type': 'text', 'index': 7},
        '日付_6日目': {'type': 'text', 'index': 8},
        '日付_7日目': {'type': 'text', 'index': 9},
        '日付_8日目': {'type': 'text', 'index': 10},
        '日付_9日目': {'type': 'text', 'index': 11},
        '日付_10日目': {'type': 'text', 'index': 12},
        '日付_11日目': {'type': 'text', 'index': 13},
        '日付_12日目': {'type': 'text', 'index': 14},
        '日付_13日目': {'type': 'text', 'index': 15},
        '日付_14日目': {'type': 'text', 'index': 16},
        '日付_15日目': {'type': 'text', 'index': 17},
        '日付_16日目': {'type': 'text', 'index': 18},
        '日付_17日目': {'type': 'text', 'index': 19},
        '日付_18日目': {'type': 'text', 'index': 20},
        '日付_19日目': {'type': 'text', 'index': 21},
        '日付_20日目': {'type': 'text', 'index': 22},
        '日付_21日目': {'type': 'text', 'index': 23},
        '日付_22日目': {'type': 'text', 'index': 24},
        '日付_23日目': {'type': 'text', 'index': 25},
        '日付_24日目': {'type': 'text', 'index': 26},
        '日付_25日目': {'type': 'text', 'index': 27},
        '日付_26日目': {'type': 'text', 'index': 28},
        '日付_27日目': {'type': 'text', 'index': 29},
        '日付_28日目': {'type': 'text', 'index': 30},
        '日付_29日目': {'type': 'text', 'index': 31},
        '日付_30日目': {'type': 'text', 'index': 32},
        '日付_31日目': {'type': 'text', 'index': 33},
        '日付_32日目': {'type': 'text', 'index': 34},
        '日付_33日目': {'type': 'text', 'index': 35},
        '日付_34日目': {'type': 'text', 'index': 36},
        '日付_35日目': {'type': 'text', 'index': 37},
        '日付_36日目': {'type': 'text', 'index': 38},
        '日付_37日目': {'type': 'text', 'index': 39},
        '日付_38日目': {'type': 'text', 'index': 40},
        '日付_39日目': {'type': 'text', 'index': 41},
        '日付_40日目': {'type': 'text', 'index': 42},
        '日付_41日目': {'type': 'text', 'index': 43},
        '日付_42日目': {'type': 'text', 'index': 44},
        '日付_43日目': {'type': 'text', 'index': 45},
        '日付_44日目': {'type': 'text', 'index': 46},
        '日付_45日目': {'type': 'text', 'index': 47},
        '日付_46日目': {'type': 'text', 'index': 48},
        '日付_47日目': {'type': 'text', 'index': 49},
        '日付_48日目': {'type': 'text', 'index': 50},
        '日付_49日目': {'type': 'text', 'index': 51},
        '日付_50日目': {'type': 'text', 'index': 52},
        '日付_51日目': {'type': 'text', 'index': 53},
        '日付_52日目': {'type': 'text', 'index': 54},
        '日付_53日目': {'type': 'text', 'index': 55},
        '日付_54日目': {'type': 'text', 'index': 56},
        '日付_55日目': {'type': 'text', 'index': 57},
        '日付_56日目': {'type': 'text', 'index': 58},
        '日付_57日目': {'type': 'text', 'index': 59},
        '日付_58日目': {'type': 'text', 'index': 60},
        '日付_59日目': {'type': 'text', 'index': 61},
        '日付_60日目': {'type': 'text', 'index': 62},
        '日付_61日目': {'type': 'text', 'index': 63},
        '日付_62日目': {'type': 'text', 'index': 64},
        '日付_63日目': {'type': 'text', 'index': 65},
        '日付_64日目': {'type': 'text', 'index': 66},
        '日付_65日目': {'type': 'text', 'index': 67},
        '日付_66日目': {'type': 'text', 'index': 68},
        '日付_67日目': {'type': 'text', 'index': 69},
        '日付_68日目': {'type': 'text', 'index': 70},
        '日付_69日目': {'type': 'text', 'index': 71},
        '日付_70日目': {'type': 'text', 'index': 72},
        '日付_71日目': {'type': 'text', 'index': 73},
        '日付_72日目': {'type': 'text', 'index': 74},
        '日付_73日目': {'type': 'text', 'index': 75},
        '日付_74日目': {'type': 'text', 'index': 76},
        '日付_75日目': {'type': 'text', 'index': 77},
        '日付_76日目': {'type': 'text', 'index': 78},
        '日付_77日目': {'type': 'text', 'index': 79},
        '日付_78日目': {'type': 'text', 'index': 80},
        '日付_79日目': {'type': 'text', 'index': 81},
        '日付_80日目': {'type': 'text', 'index': 82},
        '日付_81日目': {'type': 'text', 'index': 83},
        '日付_82日目': {'type': 'text', 'index': 84},
        '日付_83日目': {'type': 'text', 'index': 85},
        '日付_84日目': {'type': 'text', 'index': 86},
        '日付_85日目': {'type': 'text', 'index': 87},
        '日付_86日目': {'type': 'text', 'index': 88},
        '日付_87日目': {'type': 'text', 'index': 89},
        '日付_88日目': {'type': 'text', 'index': 90},
        '日付_89日目': {'type': 'text', 'index': 91},
        '日付_90日目': {'type': 'text', 'index': 92},
        '日付_91日目': {'type': 'text', 'index': 93},
        '日付_92日目': {'type': 'text', 'index': 94},
        '日付_93日目': {'type': 'text', 'index': 95},
        '日付_94日目': {'type': 'text', 'index': 96},
        '日付_95日目': {'type': 'text', 'index': 97},
        '日付_96日目': {'type': 'text', 'index': 98},
        '日付_97日目': {'type': 'text', 'index': 99},
        '日付_98日目': {'type': 'text', 'index': 100},
        '日付_99日目': {'type': 'text', 'index': 101},
        '日付_100日目': {'type': 'text', 'index': 102},
        '日付_101日目': {'type': 'text', 'index': 103},
        '日付_102日目': {'type': 'text', 'index': 104},
        '日付_103日目': {'type': 'text', 'index': 105},
        '日付_104日目': {'type': 'text', 'index': 106},
        '日付_105日目': {'type': 'text', 'index': 107},
        '日付_106日目': {'type': 'text', 'index': 108},
        '日付_107日目': {'type': 'text', 'index': 109},
        '日付_108日目': {'type': 'text', 'index': 110},
        '日付_109日目': {'type': 'text', 'index': 111},
        '日付_110日目': {'type': 'text', 'index': 112},
        '日付_111日目': {'type': 'text', 'index': 113},
        '日付_112日目': {'type': 'text', 'index': 114},
        '日付_113日目': {'type': 'text', 'index': 115},
        '日付_114日目': {'type': 'text', 'index': 116},
        '日付_115日目': {'type': 'text', 'index': 117},
        '日付_116日目': {'type': 'text', 'index': 118},
        '日付_117日目': {'type': 'text', 'index': 119},
        '日付_118日目': {'type': 'text', 'index': 120},
        '日付_119日目': {'type': 'text', 'index': 121},
        '日付_120日目': {'type': 'text', 'index': 122},
        '日付_121日目': {'type': 'text', 'index': 123},
        '日付_122日目': {'type': 'text', 'index': 124},
        '日付_123日目': {'type': 'text', 'index': 125},
        '日付_124日目': {'type': 'text', 'index': 126},
        '日付_125日目': {'type': 'text', 'index': 127},
        '日付_126日目': {'type': 'text', 'index': 128},
        '日付_127日目': {'type': 'text', 'index': 129},
        '日付_128日目': {'type': 'text', 'index': 130},
        '日付_129日目': {'type': 'text', 'index': 131},
        '日付_130日目': {'type': 'text', 'index': 132},
        '日付_131日目': {'type': 'text', 'index': 133},
        '日付_132日目': {'type': 'text', 'index': 134},
        '日付_133日目': {'type': 'text', 'index': 135},
        '日付_134日目': {'type': 'text', 'index': 136},
        '日付_135日目': {'type': 'text', 'index': 137},
        '日付_136日目': {'type': 'text', 'index': 138},
        '日付_137日目': {'type': 'text', 'index': 139},
        '日付_138日目': {'type': 'text', 'index': 140},
        '日付_139日目': {'type': 'text', 'index': 141},
        '日付_140日目': {'type': 'text', 'index': 142},
        '日付_141日目': {'type': 'text', 'index': 143},
        '日付_142日目': {'type': 'text', 'index': 144},
        '日付_143日目': {'type': 'text', 'index': 145},
        '日付_144日目': {'type': 'text', 'index': 146},
        '日付_145日目': {'type': 'text', 'index': 147},
        '日付_146日目': {'type': 'text', 'index': 148},
        '日付_147日目': {'type': 'text', 'index': 149},
        '日付_148日目': {'type': 'text', 'index': 150},
        '日付_149日目': {'type': 'text', 'index': 151},
        '日付_150日目': {'type': 'text', 'index': 152},
        '日付_151日目': {'type': 'text', 'index': 153},
        '日付_152日目': {'type': 'text', 'index': 154},
        '日付_153日目': {'type': 'text', 'index': 155},
        '日付_154日目': {'type': 'text', 'index': 156},
        '日付_155日目': {'type': 'text', 'index': 157},
        '日付_156日目': {'type': 'text', 'index': 158},
        '日付_157日目': {'type': 'text', 'index': 159},
        '日付_158日目': {'type': 'text', 'index': 160},
        '日付_159日目': {'type': 'text', 'index': 161},
        '日付_160日目': {'type': 'text', 'index': 162},
        '日付_161日目': {'type': 'text', 'index': 163},
        '日付_162日目': {'type': 'text', 'index': 164},
        '日付_163日目': {'type': 'text', 'index': 165},
        '日付_164日目': {'type': 'text', 'index': 166},
        '日付_165日目': {'type': 'text', 'index': 167},
        '日付_166日目': {'type': 'text', 'index': 168},
        '日付_167日目': {'type': 'text', 'index': 169},
        '日付_168日目': {'type': 'text', 'index': 170},
        '日付_169日目': {'type': 'text', 'index': 171},
        '日付_170日目': {'type': 'text', 'index': 172},
        '日付_171日目': {'type': 'text', 'index': 173},
        '日付_172日目': {'type': 'text', 'index': 174},
        '日付_173日目': {'type': 'text', 'index': 175},
        '日付_174日目': {'type': 'text', 'index': 176},
        '日付_175日目': {'type': 'text', 'index': 177},
        '日付_176日目': {'type': 'text', 'index': 178},
        '日付_177日目': {'type': 'text', 'index': 179},
        '日付_178日目': {'type': 'text', 'index': 180},
        '日付_179日目': {'type': 'text', 'index': 181},
        '日付_180日目': {'type': 'text', 'index': 182},
        '日付_181日目': {'type': 'text', 'index': 183},
        '日付_182日目': {'type': 'text', 'index': 184},
        '日付_183日目': {'type': 'text', 'index': 185},
        '日付_184日目': {'type': 'text', 'index': 186},
        '日付_185日目': {'type': 'text', 'index': 187},
        '日付_186日目': {'type': 'text', 'index': 188},
        '日付_187日目': {'type': 'text', 'index': 189},
        '日付_188日目': {'type': 'text', 'index': 190},
        '日付_189日目': {'type': 'text', 'index': 191},
        '日付_190日目': {'type': 'text', 'index': 192},
        '日付_191日目': {'type': 'text', 'index': 193},
        '日付_192日目': {'type': 'text', 'index': 194},
        '日付_193日目': {'type': 'text', 'index': 195},
        '日付_194日目': {'type': 'text', 'index': 196},
        '日付_195日目': {'type': 'text', 'index': 197},
        '日付_196日目': {'type': 'text', 'index': 198},
        '日付_197日目': {'type': 'text', 'index': 199},
        '日付_198日目': {'type': 'text', 'index': 200},
        '日付_199日目': {'type': 'text', 'index': 201},
        '日付_200日目': {'type': 'text', 'index': 202},
        '日付_201日目': {'type': 'text', 'index': 203},
        '日付_202日目': {'type': 'text', 'index': 204},
        '日付_203日目': {'type': 'text', 'index': 205},
        '日付_204日目': {'type': 'text', 'index': 206},
        '日付_205日目': {'type': 'text', 'index': 207},
        '日付_206日目': {'type': 'text', 'index': 208},
        '日付_207日目': {'type': 'text', 'index': 209},
        '日付_208日目': {'type': 'text', 'index': 210},
        '日付_209日目': {'type': 'text', 'index': 211},
        '日付_210日目': {'type': 'text', 'index': 212},
        '日付_211日目': {'type': 'text', 'index': 213},
        '日付_212日目': {'type': 'text', 'index': 214},
        '日付_213日目': {'type': 'text', 'index': 215},
        '日付_214日目': {'type': 'text', 'index': 216},
        '日付_215日目': {'type': 'text', 'index': 217},
        '日付_216日目': {'type': 'text', 'index': 218},
        '日付_217日目': {'type': 'text', 'index': 219},
        '日付_218日目': {'type': 'text', 'index': 220},
        '日付_219日目': {'type': 'text', 'index': 221},
        '日付_220日目': {'type': 'text', 'index': 222},
        '日付_221日目': {'type': 'text', 'index': 223},
        '日付_222日目': {'type': 'text', 'index': 224},
        '日付_223日目': {'type': 'text', 'index': 225},
        '日付_224日目': {'type': 'text', 'index': 226},
        '日付_225日目': {'type': 'text', 'index': 227},
        '日付_226日目': {'type': 'text', 'index': 228},
        '日付_227日目': {'type': 'text', 'index': 229},
        '日付_228日目': {'type': 'text', 'index': 230},
        '日付_229日目': {'type': 'text', 'index': 231},
        '日付_230日目': {'type': 'text', 'index': 232},
        '内容_1日目': {'type': 'text', 'index': 233},
        '内容_2日目': {'type': 'text', 'index': 234},
        '内容_3日目': {'type': 'text', 'index': 235},
        '内容_4日目': {'type': 'text', 'index': 236},
        '内容_5日目': {'type': 'text', 'index': 237},
        '内容_6日目': {'type': 'text', 'index': 238},
        '内容_7日目': {'type': 'text', 'index': 239},
        '内容_8日目': {'type': 'text', 'index': 240},
        '内容_9日目': {'type': 'text', 'index': 241},
        '内容_10日目': {'type': 'text', 'index': 242},
        '内容_11日目': {'type': 'text', 'index': 243},
        '内容_12日目': {'type': 'text', 'index': 244},
        '内容_13日目': {'type': 'text', 'index': 245},
        '内容_14日目': {'type': 'text', 'index': 246},
        '内容_15日目': {'type': 'text', 'index': 247},
        '内容_16日目': {'type': 'text', 'index': 248},
        '内容_17日目': {'type': 'text', 'index': 249},
        '内容_18日目': {'type': 'text', 'index': 250},
        '内容_19日目': {'type': 'text', 'index': 251},
        '内容_20日目': {'type': 'text', 'index': 252},
        '内容_21日目': {'type': 'text', 'index': 253},
        '内容_22日目': {'type': 'text', 'index': 254},
        '内容_23日目': {'type': 'text', 'index': 255},
        '内容_24日目': {'type': 'text', 'index': 256},
        '内容_25日目': {'type': 'text', 'index': 257},
        '内容_26日目': {'type': 'text', 'index': 258},
        '内容_27日目': {'type': 'text', 'index': 259},
        '内容_28日目': {'type': 'text', 'index': 260},
        '内容_29日目': {'type': 'text', 'index': 261},
        '内容_30日目': {'type': 'text', 'index': 262},
        '内容_31日目': {'type': 'text', 'index': 263},
        '内容_32日目': {'type': 'text', 'index': 264},
        '内容_33日目': {'type': 'text', 'index': 265},
        '内容_34日目': {'type': 'text', 'index': 266},
        '内容_35日目': {'type': 'text', 'index': 267},
        '内容_36日目': {'type': 'text', 'index': 268},
        '内容_37日目': {'type': 'text', 'index': 269},
        '内容_38日目': {'type': 'text', 'index': 270},
        '内容_39日目': {'type': 'text', 'index': 271},
        '内容_40日目': {'type': 'text', 'index': 272},
        '内容_41日目': {'type': 'text', 'index': 273},
        '内容_42日目': {'type': 'text', 'index': 274},
        '内容_43日目': {'type': 'text', 'index': 275},
        '内容_44日目': {'type': 'text', 'index': 276},
        '内容_45日目': {'type': 'text', 'index': 277},
        '内容_46日目': {'type': 'text', 'index': 278},
        '内容_47日目': {'type': 'text', 'index': 279},
        '内容_48日目': {'type': 'text', 'index': 280},
        '内容_49日目': {'type': 'text', 'index': 281},
        '内容_50日目': {'type': 'text', 'index': 282},
        '内容_51日目': {'type': 'text', 'index': 283},
        '内容_52日目': {'type': 'text', 'index': 284},
        '内容_53日目': {'type': 'text', 'index': 285},
        '内容_54日目': {'type': 'text', 'index': 286},
        '内容_55日目': {'type': 'text', 'index': 287},
        '内容_56日目': {'type': 'text', 'index': 288},
        '内容_57日目': {'type': 'text', 'index': 289},
        '内容_58日目': {'type': 'text', 'index': 290},
        '内容_59日目': {'type': 'text', 'index': 291},
        '内容_60日目': {'type': 'text', 'index': 292},
        '内容_61日目': {'type': 'text', 'index': 293},
        '内容_62日目': {'type': 'text', 'index': 294},
        '内容_63日目': {'type': 'text', 'index': 295},
        '内容_64日目': {'type': 'text', 'index': 296},
        '内容_65日目': {'type': 'text', 'index': 297},
        '内容_66日目': {'type': 'text', 'index': 298},
        '内容_67日目': {'type': 'text', 'index': 299},
        '内容_68日目': {'type': 'text', 'index': 300},
        '内容_69日目': {'type': 'text', 'index': 301},
        '内容_70日目': {'type': 'text', 'index': 302},
        '内容_71日目': {'type': 'text', 'index': 303},
        '内容_72日目': {'type': 'text', 'index': 304},
        '内容_73日目': {'type': 'text', 'index': 305},
        '内容_74日目': {'type': 'text', 'index': 306},
        '内容_75日目': {'type': 'text', 'index': 307},
        '内容_76日目': {'type': 'text', 'index': 308},
        '内容_77日目': {'type': 'text', 'index': 309},
        '内容_78日目': {'type': 'text', 'index': 310},
        '内容_79日目': {'type': 'text', 'index': 311},
        '内容_80日目': {'type': 'text', 'index': 312},
        '内容_81日目': {'type': 'text', 'index': 313},
        '内容_82日目': {'type': 'text', 'index': 314},
        '内容_83日目': {'type': 'text', 'index': 315},
        '内容_84日目': {'type': 'text', 'index': 316},
        '内容_85日目': {'type': 'text', 'index': 317},
        '内容_86日目': {'type': 'text', 'index': 318},
        '内容_87日目': {'type': 'text', 'index': 319},
        '内容_88日目': {'type': 'text', 'index': 320},
        '内容_89日目': {'type': 'text', 'index': 321},
        '内容_90日目': {'type': 'text', 'index': 322},
        '内容_91日目': {'type': 'text', 'index': 323},
        '内容_92日目': {'type': 'text', 'index': 324},
        '内容_93日目': {'type': 'text', 'index': 325},
        '内容_94日目': {'type': 'text', 'index': 326},
        '内容_95日目': {'type': 'text', 'index': 327},
        '内容_96日目': {'type': 'text', 'index': 328},
        '内容_97日目': {'type': 'text', 'index': 329},
        '内容_98日目': {'type': 'text', 'index': 330},
        '内容_99日目': {'type': 'text', 'index': 331},
        '内容_100日目': {'type': 'text', 'index': 332},
        '内容_101日目': {'type': 'text', 'index': 333},
        '内容_102日目': {'type': 'text', 'index': 334},
        '内容_103日目': {'type': 'text', 'index': 335},
        '内容_104日目': {'type': 'text', 'index': 336},
        '内容_105日目': {'type': 'text', 'index': 337},
        '内容_106日目': {'type': 'text', 'index': 338},
        '内容_107日目': {'type': 'text', 'index': 339},
        '内容_108日目': {'type': 'text', 'index': 340},
        '内容_109日目': {'type': 'text', 'index': 341},
        '内容_110日目': {'type': 'text', 'index': 342},
        '内容_111日目': {'type': 'text', 'index': 343},
        '内容_112日目': {'type': 'text', 'index': 344},
        '内容_113日目': {'type': 'text', 'index': 345},
        '内容_114日目': {'type': 'text', 'index': 346},
        '内容_115日目': {'type': 'text', 'index': 347},
        '内容_116日目': {'type': 'text', 'index': 348},
        '内容_117日目': {'type': 'text', 'index': 349},
        '内容_118日目': {'type': 'text', 'index': 350},
        '内容_119日目': {'type': 'text', 'index': 351},
        '内容_120日目': {'type': 'text', 'index': 352},
        '内容_121日目': {'type': 'text', 'index': 353},
        '内容_122日目': {'type': 'text', 'index': 354},
        '内容_123日目': {'type': 'text', 'index': 355},
        '内容_124日目': {'type': 'text', 'index': 356},
        '内容_125日目': {'type': 'text', 'index': 357},
        '内容_126日目': {'type': 'text', 'index': 358},
        '内容_127日目': {'type': 'text', 'index': 359},
        '内容_128日目': {'type': 'text', 'index': 360},
        '内容_129日目': {'type': 'text', 'index': 361},
        '内容_130日目': {'type': 'text', 'index': 362},
        '内容_131日目': {'type': 'text', 'index': 363},
        '内容_132日目': {'type': 'text', 'index': 364},
        '内容_133日目': {'type': 'text', 'index': 365},
        '内容_134日目': {'type': 'text', 'index': 366},
        '内容_135日目': {'type': 'text', 'index': 367},
        '内容_136日目': {'type': 'text', 'index': 368},
        '内容_137日目': {'type': 'text', 'index': 369},
        '内容_138日目': {'type': 'text', 'index': 370},
        '内容_139日目': {'type': 'text', 'index': 371},
        '内容_140日目': {'type': 'text', 'index': 372},
        '内容_141日目': {'type': 'text', 'index': 373},
        '内容_142日目': {'type': 'text', 'index': 374},
        '内容_143日目': {'type': 'text', 'index': 375},
        '内容_144日目': {'type': 'text', 'index': 376},
        '内容_145日目': {'type': 'text', 'index': 377},
        '内容_146日目': {'type': 'text', 'index': 378},
        '内容_147日目': {'type': 'text', 'index': 379},
        '内容_148日目': {'type': 'text', 'index': 380},
        '内容_149日目': {'type': 'text', 'index': 381},
        '内容_150日目': {'type': 'text', 'index': 382},
        '内容_151日目': {'type': 'text', 'index': 383},
        '内容_152日目': {'type': 'text', 'index': 384},
        '内容_153日目': {'type': 'text', 'index': 385},
        '内容_154日目': {'type': 'text', 'index': 386},
        '内容_155日目': {'type': 'text', 'index': 387},
        '内容_156日目': {'type': 'text', 'index': 388},
        '内容_157日目': {'type': 'text', 'index': 389},
        '内容_158日目': {'type': 'text', 'index': 390},
        '内容_159日目': {'type': 'text', 'index': 391},
        '内容_160日目': {'type': 'text', 'index': 392},
        '内容_161日目': {'type': 'text', 'index': 393},
        '内容_162日目': {'type': 'text', 'index': 394},
        '内容_163日目': {'type': 'text', 'index': 395},
        '内容_164日目': {'type': 'text', 'index': 396},
        '内容_165日目': {'type': 'text', 'index': 397},
        '内容_166日目': {'type': 'text', 'index': 398},
        '内容_167日目': {'type': 'text', 'index': 399},
        '内容_168日目': {'type': 'text', 'index': 400},
        '内容_169日目': {'type': 'text', 'index': 401},
        '内容_170日目': {'type': 'text', 'index': 402},
        '内容_171日目': {'type': 'text', 'index': 403},
        '内容_172日目': {'type': 'text', 'index': 404},
        '内容_173日目': {'type': 'text', 'index': 405},
        '内容_174日目': {'type': 'text', 'index': 406},
        '内容_175日目': {'type': 'text', 'index': 407},
        '内容_176日目': {'type': 'text', 'index': 408},
        '内容_177日目': {'type': 'text', 'index': 409},
        '内容_178日目': {'type': 'text', 'index': 410},
        '内容_179日目': {'type': 'text', 'index': 411},
        '内容_180日目': {'type': 'text', 'index': 412},
        '内容_181日目': {'type': 'text', 'index': 413},
        '内容_182日目': {'type': 'text', 'index': 414},
        '内容_183日目': {'type': 'text', 'index': 415},
        '内容_184日目': {'type': 'text', 'index': 416},
        '内容_185日目': {'type': 'text', 'index': 417},
        '内容_186日目': {'type': 'text', 'index': 418},
        '内容_187日目': {'type': 'text', 'index': 419},
        '内容_188日目': {'type': 'text', 'index': 420},
        '内容_189日目': {'type': 'text', 'index': 421},
        '内容_190日目': {'type': 'text', 'index': 422},
        '内容_191日目': {'type': 'text', 'index': 423},
        '内容_192日目': {'type': 'text', 'index': 424},
        '内容_193日目': {'type': 'text', 'index': 425},
        '内容_194日目': {'type': 'text', 'index': 426},
        '内容_195日目': {'type': 'text', 'index': 427},
        '内容_196日目': {'type': 'text', 'index': 428},
        '内容_197日目': {'type': 'text', 'index': 429},
        '内容_198日目': {'type': 'text', 'index': 430},
        '内容_199日目': {'type': 'text', 'index': 431},
        '内容_200日目': {'type': 'text', 'index': 432},
        '内容_201日目': {'type': 'text', 'index': 433},
        '内容_202日目': {'type': 'text', 'index': 434},
        '内容_203日目': {'type': 'text', 'index': 435},
        '内容_204日目': {'type': 'text', 'index': 436},
        '内容_205日目': {'type': 'text', 'index': 437},
        '内容_206日目': {'type': 'text', 'index': 438},
        '内容_207日目': {'type': 'text', 'index': 439},
        '内容_208日目': {'type': 'text', 'index': 440},
        '内容_209日目': {'type': 'text', 'index': 441},
        '内容_210日目': {'type': 'text', 'index': 442},
        '内容_211日目': {'type': 'text', 'index': 443},
        '内容_212日目': {'type': 'text', 'index': 444},
        '内容_213日目': {'type': 'text', 'index': 445},
        '内容_214日目': {'type': 'text', 'index': 446},
        '内容_215日目': {'type': 'text', 'index': 447},
        '内容_216日目': {'type': 'text', 'index': 448},
        '内容_217日目': {'type': 'text', 'index': 449},
        '内容_218日目': {'type': 'text', 'index': 450},
        '内容_219日目': {'type': 'text', 'index': 451},
        '内容_220日目': {'type': 'text', 'index': 452},
        '内容_221日目': {'type': 'text', 'index': 453},
        '内容_222日目': {'type': 'text', 'index': 454},
        '内容_223日目': {'type': 'text', 'index': 455},
        '内容_224日目': {'type': 'text', 'index': 456},
        '内容_225日目': {'type': 'text', 'index': 457},
        '内容_226日目': {'type': 'text', 'index': 458},
        '内容_227日目': {'type': 'text', 'index': 459},
        '内容_228日目': {'type': 'text', 'index': 460},
        '内容_229日目': {'type': 'text', 'index': 461},
        '内容_230日目': {'type': 'text', 'index': 462},
        '担当_1日目': {'type': 'text', 'index': 463},
        '担当_2日目': {'type': 'text', 'index': 464},
        '担当_3日目': {'type': 'text', 'index': 465},
        '担当_4日目': {'type': 'text', 'index': 466},
        '担当_5日目': {'type': 'text', 'index': 467},
        '担当_6日目': {'type': 'text', 'index': 468},
        '担当_7日目': {'type': 'text', 'index': 469},
        '担当_8日目': {'type': 'text', 'index': 470},
        '担当_9日目': {'type': 'text', 'index': 471},
        '担当_10日目': {'type': 'text', 'index': 472},
        '担当_11日目': {'type': 'text', 'index': 473},
        '担当_12日目': {'type': 'text', 'index': 474},
        '担当_13日目': {'type': 'text', 'index': 475},
        '担当_14日目': {'type': 'text', 'index': 476},
        '担当_15日目': {'type': 'text', 'index': 477},
        '担当_16日目': {'type': 'text', 'index': 478},
        '担当_17日目': {'type': 'text', 'index': 479},
        '担当_18日目': {'type': 'text', 'index': 480},
        '担当_19日目': {'type': 'text', 'index': 481},
        '担当_20日目': {'type': 'text', 'index': 482},
        '担当_21日目': {'type': 'text', 'index': 483},
        '担当_22日目': {'type': 'text', 'index': 484},
        '担当_23日目': {'type': 'text', 'index': 485},
        '担当_24日目': {'type': 'text', 'index': 486},
        '担当_25日目': {'type': 'text', 'index': 487},
        '担当_26日目': {'type': 'text', 'index': 488},
        '担当_27日目': {'type': 'text', 'index': 489},
        '担当_28日目': {'type': 'text', 'index': 490},
        '担当_29日目': {'type': 'text', 'index': 491},
        '担当_30日目': {'type': 'text', 'index': 492},
        '担当_31日目': {'type': 'text', 'index': 493},
        '担当_32日目': {'type': 'text', 'index': 494},
        '担当_33日目': {'type': 'text', 'index': 495},
        '担当_34日目': {'type': 'text', 'index': 496},
        '担当_35日目': {'type': 'text', 'index': 497},
        '担当_36日目': {'type': 'text', 'index': 498},
        '担当_37日目': {'type': 'text', 'index': 499},
        '担当_38日目': {'type': 'text', 'index': 500},
        '担当_39日目': {'type': 'text', 'index': 501},
        '担当_40日目': {'type': 'text', 'index': 502},
        '担当_41日目': {'type': 'text', 'index': 503},
        '担当_42日目': {'type': 'text', 'index': 504},
        '担当_43日目': {'type': 'text', 'index': 505},
        '担当_44日目': {'type': 'text', 'index': 506},
        '担当_45日目': {'type': 'text', 'index': 507},
        '担当_46日目': {'type': 'text', 'index': 508},
        '担当_47日目': {'type': 'text', 'index': 509},
        '担当_48日目': {'type': 'text', 'index': 510},
        '担当_49日目': {'type': 'text', 'index': 511},
        '担当_50日目': {'type': 'text', 'index': 512},
        '担当_51日目': {'type': 'text', 'index': 513},
        '担当_52日目': {'type': 'text', 'index': 514},
        '担当_53日目': {'type': 'text', 'index': 515},
        '担当_54日目': {'type': 'text', 'index': 516},
        '担当_55日目': {'type': 'text', 'index': 517},
        '担当_56日目': {'type': 'text', 'index': 518},
        '担当_57日目': {'type': 'text', 'index': 519},
        '担当_58日目': {'type': 'text', 'index': 520},
        '担当_59日目': {'type': 'text', 'index': 521},
        '担当_60日目': {'type': 'text', 'index': 522},
        '担当_61日目': {'type': 'text', 'index': 523},
        '担当_62日目': {'type': 'text', 'index': 524},
        '担当_63日目': {'type': 'text', 'index': 525},
        '担当_64日目': {'type': 'text', 'index': 526},
        '担当_65日目': {'type': 'text', 'index': 527},
        '担当_66日目': {'type': 'text', 'index': 528},
        '担当_67日目': {'type': 'text', 'index': 529},
        '担当_68日目': {'type': 'text', 'index': 530},
        '担当_69日目': {'type': 'text', 'index': 531},
        '担当_70日目': {'type': 'text', 'index': 532},
        '担当_71日目': {'type': 'text', 'index': 533},
        '担当_72日目': {'type': 'text', 'index': 534},
        '担当_73日目': {'type': 'text', 'index': 535},
        '担当_74日目': {'type': 'text', 'index': 536},
        '担当_75日目': {'type': 'text', 'index': 537},
        '担当_76日目': {'type': 'text', 'index': 538},
        '担当_77日目': {'type': 'text', 'index': 539},
        '担当_78日目': {'type': 'text', 'index': 540},
        '担当_79日目': {'type': 'text', 'index': 541},
        '担当_80日目': {'type': 'text', 'index': 542},
        '担当_81日目': {'type': 'text', 'index': 543},
        '担当_82日目': {'type': 'text', 'index': 544},
        '担当_83日目': {'type': 'text', 'index': 545},
        '担当_84日目': {'type': 'text', 'index': 546},
        '担当_85日目': {'type': 'text', 'index': 547},
        '担当_86日目': {'type': 'text', 'index': 548},
        '担当_87日目': {'type': 'text', 'index': 549},
        '担当_88日目': {'type': 'text', 'index': 550},
        '担当_89日目': {'type': 'text', 'index': 551},
        '担当_90日目': {'type': 'text', 'index': 552},
        '担当_91日目': {'type': 'text', 'index': 553},
        '担当_92日目': {'type': 'text', 'index': 554},
        '担当_93日目': {'type': 'text', 'index': 555},
        '担当_94日目': {'type': 'text', 'index': 556},
        '担当_95日目': {'type': 'text', 'index': 557},
        '担当_96日目': {'type': 'text', 'index': 558},
        '担当_97日目': {'type': 'text', 'index': 559},
        '担当_98日目': {'type': 'text', 'index': 560},
        '担当_99日目': {'type': 'text', 'index': 561},
        '担当_100日目': {'type': 'text', 'index': 562},
        '担当_101日目': {'type': 'text', 'index': 563},
        '担当_102日目': {'type': 'text', 'index': 564},
        '担当_103日目': {'type': 'text', 'index': 565},
        '担当_104日目': {'type': 'text', 'index': 566},
        '担当_105日目': {'type': 'text', 'index': 567},
        '担当_106日目': {'type': 'text', 'index': 568},
        '担当_107日目': {'type': 'text', 'index': 569},
        '担当_108日目': {'type': 'text', 'index': 570},
        '担当_109日目': {'type': 'text', 'index': 571},
        '担当_110日目': {'type': 'text', 'index': 572},
        '担当_111日目': {'type': 'text', 'index': 573},
        '担当_112日目': {'type': 'text', 'index': 574},
        '担当_113日目': {'type': 'text', 'index': 575},
        '担当_114日目': {'type': 'text', 'index': 576},
        '担当_115日目': {'type': 'text', 'index': 577},
        '担当_116日目': {'type': 'text', 'index': 578},
        '担当_117日目': {'type': 'text', 'index': 579},
        '担当_118日目': {'type': 'text', 'index': 580},
        '担当_119日目': {'type': 'text', 'index': 581},
        '担当_120日目': {'type': 'text', 'index': 582},
        '担当_121日目': {'type': 'text', 'index': 583},
        '担当_122日目': {'type': 'text', 'index': 584},
        '担当_123日目': {'type': 'text', 'index': 585},
        '担当_124日目': {'type': 'text', 'index': 586},
        '担当_125日目': {'type': 'text', 'index': 587},
        '担当_126日目': {'type': 'text', 'index': 588},
        '担当_127日目': {'type': 'text', 'index': 589},
        '担当_128日目': {'type': 'text', 'index': 590},
        '担当_129日目': {'type': 'text', 'index': 591},
        '担当_130日目': {'type': 'text', 'index': 592},
        '担当_131日目': {'type': 'text', 'index': 593},
        '担当_132日目': {'type': 'text', 'index': 594},
        '担当_133日目': {'type': 'text', 'index': 595},
        '担当_134日目': {'type': 'text', 'index': 596},
        '担当_135日目': {'type': 'text', 'index': 597},
        '担当_136日目': {'type': 'text', 'index': 598},
        '担当_137日目': {'type': 'text', 'index': 599},
        '担当_138日目': {'type': 'text', 'index': 600},
        '担当_139日目': {'type': 'text', 'index': 601},
        '担当_140日目': {'type': 'text', 'index': 602},
        '担当_141日目': {'type': 'text', 'index': 603},
        '担当_142日目': {'type': 'text', 'index': 604},
        '担当_143日目': {'type': 'text', 'index': 605},
        '担当_144日目': {'type': 'text', 'index': 606},
        '担当_145日目': {'type': 'text', 'index': 607},
        '担当_146日目': {'type': 'text', 'index': 608},
        '担当_147日目': {'type': 'text', 'index': 609},
        '担当_148日目': {'type': 'text', 'index': 610},
        '担当_149日目': {'type': 'text', 'index': 611},
        '担当_150日目': {'type': 'text', 'index': 612},
        '担当_151日目': {'type': 'text', 'index': 613},
        '担当_152日目': {'type': 'text', 'index': 614},
        '担当_153日目': {'type': 'text', 'index': 615},
        '担当_154日目': {'type': 'text', 'index': 616},
        '担当_155日目': {'type': 'text', 'index': 617},
        '担当_156日目': {'type': 'text', 'index': 618},
        '担当_157日目': {'type': 'text', 'index': 619},
        '担当_158日目': {'type': 'text', 'index': 620},
        '担当_159日目': {'type': 'text', 'index': 621},
        '担当_160日目': {'type': 'text', 'index': 622},
        '担当_161日目': {'type': 'text', 'index': 623},
        '担当_162日目': {'type': 'text', 'index': 624},
        '担当_163日目': {'type': 'text', 'index': 625},
        '担当_164日目': {'type': 'text', 'index': 626},
        '担当_165日目': {'type': 'text', 'index': 627},
        '担当_166日目': {'type': 'text', 'index': 628},
        '担当_167日目': {'type': 'text', 'index': 629},
        '担当_168日目': {'type': 'text', 'index': 630},
        '担当_169日目': {'type': 'text', 'index': 631},
        '担当_170日目': {'type': 'text', 'index': 632},
        '担当_171日目': {'type': 'text', 'index': 633},
        '担当_172日目': {'type': 'text', 'index': 634},
        '担当_173日目': {'type': 'text', 'index': 635},
        '担当_174日目': {'type': 'text', 'index': 636},
        '担当_175日目': {'type': 'text', 'index': 637},
        '担当_176日目': {'type': 'text', 'index': 638},
        '担当_177日目': {'type': 'text', 'index': 639},
        '担当_178日目': {'type': 'text', 'index': 640},
        '担当_179日目': {'type': 'text', 'index': 641},
        '担当_180日目': {'type': 'text', 'index': 642},
        '担当_181日目': {'type': 'text', 'index': 643},
        '担当_182日目': {'type': 'text', 'index': 644},
        '担当_183日目': {'type': 'text', 'index': 645},
        '担当_184日目': {'type': 'text', 'index': 646},
        '担当_185日目': {'type': 'text', 'index': 647},
        '担当_186日目': {'type': 'text', 'index': 648},
        '担当_187日目': {'type': 'text', 'index': 649},
        '担当_188日目': {'type': 'text', 'index': 650},
        '担当_189日目': {'type': 'text', 'index': 651},
        '担当_190日目': {'type': 'text', 'index': 652},
        '担当_191日目': {'type': 'text', 'index': 653},
        '担当_192日目': {'type': 'text', 'index': 654},
        '担当_193日目': {'type': 'text', 'index': 655},
        '担当_194日目': {'type': 'text', 'index': 656},
        '担当_195日目': {'type': 'text', 'index': 657},
        '担当_196日目': {'type': 'text', 'index': 658},
        '担当_197日目': {'type': 'text', 'index': 659},
        '担当_198日目': {'type': 'text', 'index': 660},
        '担当_199日目': {'type': 'text', 'index': 661},
        '担当_200日目': {'type': 'text', 'index': 662},
        '担当_201日目': {'type': 'text', 'index': 663},
        '担当_202日目': {'type': 'text', 'index': 664},
        '担当_203日目': {'type': 'text', 'index': 665},
        '担当_204日目': {'type': 'text', 'index': 666},
        '担当_205日目': {'type': 'text', 'index': 667},
        '担当_206日目': {'type': 'text', 'index': 668},
        '担当_207日目': {'type': 'text', 'index': 669},
        '担当_208日目': {'type': 'text', 'index': 670},
        '担当_209日目': {'type': 'text', 'index': 671},
        '担当_210日目': {'type': 'text', 'index': 672},
        '担当_211日目': {'type': 'text', 'index': 673},
        '担当_212日目': {'type': 'text', 'index': 674},
    }
}


# データフレームをデータベースへアップロード
def db_insert(db, db_map):
    # conn = sqlite3.connect('S:/個人作業用/渡邊/ワールドジャパン/sqlite3/salon_G.db')
    conn = sqlite3.connect('S:/個人作業用/渡邊/ワールドジャパン/sqlite3/salon_all.db')
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    # 関数db_initを呼び出し
    db_init(db_map, cur)
    col_name = []
    val = []
    table_name = db_map['table_name']

    for k, v in db_map['column'].items():
        col_name.append(k)
        val.append(v['index'])
    
    col_names = ','.join(col_name)

    flag = 0

    for r in db.iterrows():
        values = []
        if r[1][0] is None:
            flag += 1
        else:
            flag = 0
            for v in val:
                cell_val = r[1][v-1]
                if type(cell_val) is not str and type(cell_val) is not int and cell_val is not None:
                    values.append(str(cell_val))
                else:
                    values.append(cell_val)

            place_holder = ','.join('?'*len(values))
            sql = f'INSERT INTO {table_name} ({col_names}) VALUES({place_holder})'
            cur.execute(sql, tuple(values))
        
        if flag == 2:
            break

    conn.commit()
    cur.close()
    conn.close()


# 既存のテーブルの値を削除
def db_init(db_map, db):
    param = []
    table_name = db_map['table_name']

    for k, v in db_map['column'].items():
        param.append(f"{k} {v['type']}")

    params = ','.join(param)

    db.execute(f'CREATE TABLE IF NOT EXISTS {table_name}({params})')
    db.execute(f'DELETE FROM {table_name}')

# excelファイルの内容を読み込み
wb_1 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん - コピー/2フォーマット(潮永さん）/【G】顧客情報(来店時).xlsx', sheet_name='【G】顧客情報', header=1)
# 不足している列を追加
wb_1.insert(loc = 158,column = '日付_157日目', value = None)
wb_1.insert(loc = 159,column = '日付_158日目', value = None)
wb_1.insert(loc = 160,column = '日付_159日目', value = None)
wb_1.insert(loc = 161,column = '日付_160日目', value = None)
wb_1.insert(loc = 162,column = '日付_161日目', value = None)
wb_1.insert(loc = 163,column = '日付_162日目', value = None)
wb_1.insert(loc = 164,column = '日付_163日目', value = None)
wb_1.insert(loc = 165,column = '日付_164日目', value = None)
wb_1.insert(loc = 166,column = '日付_165日目', value = None)
wb_1.insert(loc = 167,column = '日付_166日目', value = None)
wb_1.insert(loc = 168,column = '日付_167日目', value = None)
wb_1.insert(loc = 169,column = '日付_168日目', value = None)
wb_1.insert(loc = 170,column = '日付_169日目', value = None)
wb_1.insert(loc = 171,column = '日付_170日目', value = None)
wb_1.insert(loc = 172,column = '日付_171日目', value = None)
wb_1.insert(loc = 173,column = '日付_172日目', value = None)
wb_1.insert(loc = 174,column = '日付_173日目', value = None)
wb_1.insert(loc = 175,column = '日付_174日目', value = None)
wb_1.insert(loc = 176,column = '日付_175日目', value = None)
wb_1.insert(loc = 177,column = '日付_176日目', value = None)
wb_1.insert(loc = 178,column = '日付_177日目', value = None)
wb_1.insert(loc = 179,column = '日付_178日目', value = None)
wb_1.insert(loc = 180,column = '日付_179日目', value = None)
wb_1.insert(loc = 181,column = '日付_180日目', value = None)
wb_1.insert(loc = 182,column = '日付_181日目', value = None)
wb_1.insert(loc = 183,column = '日付_182日目', value = None)
wb_1.insert(loc = 184,column = '日付_183日目', value = None)
wb_1.insert(loc = 185,column = '日付_184日目', value = None)
wb_1.insert(loc = 186,column = '日付_185日目', value = None)
wb_1.insert(loc = 187,column = '日付_186日目', value = None)
wb_1.insert(loc = 188,column = '日付_187日目', value = None)
wb_1.insert(loc = 189,column = '日付_188日目', value = None)
wb_1.insert(loc = 190,column = '日付_189日目', value = None)
wb_1.insert(loc = 191,column = '日付_190日目', value = None)
wb_1.insert(loc = 192,column = '日付_191日目', value = None)
wb_1.insert(loc = 193,column = '日付_192日目', value = None)
wb_1.insert(loc = 194,column = '日付_193日目', value = None)
wb_1.insert(loc = 195,column = '日付_194日目', value = None)
wb_1.insert(loc = 196,column = '日付_195日目', value = None)
wb_1.insert(loc = 197,column = '日付_196日目', value = None)
wb_1.insert(loc = 198,column = '日付_197日目', value = None)
wb_1.insert(loc = 199,column = '日付_198日目', value = None)
wb_1.insert(loc = 200,column = '日付_199日目', value = None)
wb_1.insert(loc = 201,column = '日付_200日目', value = None)
wb_1.insert(loc = 202,column = '日付_201日目', value = None)
wb_1.insert(loc = 203,column = '日付_202日目', value = None)
wb_1.insert(loc = 204,column = '日付_203日目', value = None)
wb_1.insert(loc = 205,column = '日付_204日目', value = None)
wb_1.insert(loc = 206,column = '日付_205日目', value = None)
wb_1.insert(loc = 207,column = '日付_206日目', value = None)
wb_1.insert(loc = 208,column = '日付_207日目', value = None)
wb_1.insert(loc = 209,column = '日付_208日目', value = None)
wb_1.insert(loc = 210,column = '日付_209日目', value = None)
wb_1.insert(loc = 211,column = '日付_210日目', value = None)
wb_1.insert(loc = 212,column = '日付_211日目', value = None)
wb_1.insert(loc = 213,column = '日付_212日目', value = None)
wb_1.insert(loc = 214,column = '日付_213日目', value = None)
wb_1.insert(loc = 215,column = '日付_214日目', value = None)
wb_1.insert(loc = 216,column = '日付_215日目', value = None)
wb_1.insert(loc = 217,column = '日付_216日目', value = None)
wb_1.insert(loc = 218,column = '日付_217日目', value = None)
wb_1.insert(loc = 219,column = '日付_218日目', value = None)
wb_1.insert(loc = 220,column = '日付_219日目', value = None)
wb_1.insert(loc = 221,column = '日付_220日目', value = None)
wb_1.insert(loc = 222,column = '日付_221日目', value = None)
wb_1.insert(loc = 223,column = '日付_222日目', value = None)
wb_1.insert(loc = 224,column = '日付_223日目', value = None)
wb_1.insert(loc = 225,column = '日付_224日目', value = None)
wb_1.insert(loc = 226,column = '日付_225日目', value = None)
wb_1.insert(loc = 227,column = '日付_226日目', value = None)
wb_1.insert(loc = 228,column = '日付_227日目', value = None)
wb_1.insert(loc = 229,column = '日付_228日目', value = None)
wb_1.insert(loc = 230,column = '日付_229日目', value = None)
wb_1.insert(loc = 231,column = '日付_230日目', value = None)
wb_1.insert(loc = 388,column = '内容_157日目', value = None)
wb_1.insert(loc = 389,column = '内容_158日目', value = None)
wb_1.insert(loc = 390,column = '内容_159日目', value = None)
wb_1.insert(loc = 391,column = '内容_160日目', value = None)
wb_1.insert(loc = 392,column = '内容_161日目', value = None)
wb_1.insert(loc = 393,column = '内容_162日目', value = None)
wb_1.insert(loc = 394,column = '内容_163日目', value = None)
wb_1.insert(loc = 395,column = '内容_164日目', value = None)
wb_1.insert(loc = 396,column = '内容_165日目', value = None)
wb_1.insert(loc = 397,column = '内容_166日目', value = None)
wb_1.insert(loc = 398,column = '内容_167日目', value = None)
wb_1.insert(loc = 399,column = '内容_168日目', value = None)
wb_1.insert(loc = 400,column = '内容_169日目', value = None)
wb_1.insert(loc = 401,column = '内容_170日目', value = None)
wb_1.insert(loc = 402,column = '内容_171日目', value = None)
wb_1.insert(loc = 403,column = '内容_172日目', value = None)
wb_1.insert(loc = 404,column = '内容_173日目', value = None)
wb_1.insert(loc = 405,column = '内容_174日目', value = None)
wb_1.insert(loc = 406,column = '内容_175日目', value = None)
wb_1.insert(loc = 407,column = '内容_176日目', value = None)
wb_1.insert(loc = 408,column = '内容_177日目', value = None)
wb_1.insert(loc = 409,column = '内容_178日目', value = None)
wb_1.insert(loc = 410,column = '内容_179日目', value = None)
wb_1.insert(loc = 411,column = '内容_180日目', value = None)
wb_1.insert(loc = 412,column = '内容_181日目', value = None)
wb_1.insert(loc = 413,column = '内容_182日目', value = None)
wb_1.insert(loc = 414,column = '内容_183日目', value = None)
wb_1.insert(loc = 415,column = '内容_184日目', value = None)
wb_1.insert(loc = 416,column = '内容_185日目', value = None)
wb_1.insert(loc = 417,column = '内容_186日目', value = None)
wb_1.insert(loc = 418,column = '内容_187日目', value = None)
wb_1.insert(loc = 419,column = '内容_188日目', value = None)
wb_1.insert(loc = 420,column = '内容_189日目', value = None)
wb_1.insert(loc = 421,column = '内容_190日目', value = None)
wb_1.insert(loc = 422,column = '内容_191日目', value = None)
wb_1.insert(loc = 423,column = '内容_192日目', value = None)
wb_1.insert(loc = 424,column = '内容_193日目', value = None)
wb_1.insert(loc = 425,column = '内容_194日目', value = None)
wb_1.insert(loc = 426,column = '内容_195日目', value = None)
wb_1.insert(loc = 427,column = '内容_196日目', value = None)
wb_1.insert(loc = 428,column = '内容_197日目', value = None)
wb_1.insert(loc = 429,column = '内容_198日目', value = None)
wb_1.insert(loc = 430,column = '内容_199日目', value = None)
wb_1.insert(loc = 431,column = '内容_200日目', value = None)
wb_1.insert(loc = 432,column = '内容_201日目', value = None)
wb_1.insert(loc = 433,column = '内容_202日目', value = None)
wb_1.insert(loc = 434,column = '内容_203日目', value = None)
wb_1.insert(loc = 435,column = '内容_204日目', value = None)
wb_1.insert(loc = 436,column = '内容_205日目', value = None)
wb_1.insert(loc = 437,column = '内容_206日目', value = None)
wb_1.insert(loc = 438,column = '内容_207日目', value = None)
wb_1.insert(loc = 439,column = '内容_208日目', value = None)
wb_1.insert(loc = 440,column = '内容_209日目', value = None)
wb_1.insert(loc = 441,column = '内容_210日目', value = None)
wb_1.insert(loc = 442,column = '内容_211日目', value = None)
wb_1.insert(loc = 443,column = '内容_212日目', value = None)
wb_1.insert(loc = 444,column = '内容_213日目', value = None)
wb_1.insert(loc = 445,column = '内容_214日目', value = None)
wb_1.insert(loc = 446,column = '内容_215日目', value = None)
wb_1.insert(loc = 447,column = '内容_216日目', value = None)
wb_1.insert(loc = 448,column = '内容_217日目', value = None)
wb_1.insert(loc = 449,column = '内容_218日目', value = None)
wb_1.insert(loc = 450,column = '内容_219日目', value = None)
wb_1.insert(loc = 451,column = '内容_220日目', value = None)
wb_1.insert(loc = 452,column = '内容_221日目', value = None)
wb_1.insert(loc = 453,column = '内容_222日目', value = None)
wb_1.insert(loc = 454,column = '内容_223日目', value = None)
wb_1.insert(loc = 455,column = '内容_224日目', value = None)
wb_1.insert(loc = 456,column = '内容_225日目', value = None)
wb_1.insert(loc = 457,column = '内容_226日目', value = None)
wb_1.insert(loc = 458,column = '内容_227日目', value = None)
wb_1.insert(loc = 459,column = '内容_228日目', value = None)
wb_1.insert(loc = 460,column = '内容_229日目', value = None)
wb_1.insert(loc = 461,column = '内容_230日目', value = None)
wb_1.insert(loc = 618,column = '担当_157日目', value = None)
wb_1.insert(loc = 619,column = '担当_158日目', value = None)
wb_1.insert(loc = 620,column = '担当_159日目', value = None)
wb_1.insert(loc = 621,column = '担当_160日目', value = None)
wb_1.insert(loc = 622,column = '担当_161日目', value = None)
wb_1.insert(loc = 623,column = '担当_162日目', value = None)
wb_1.insert(loc = 624,column = '担当_163日目', value = None)
wb_1.insert(loc = 625,column = '担当_164日目', value = None)
wb_1.insert(loc = 626,column = '担当_165日目', value = None)
wb_1.insert(loc = 627,column = '担当_166日目', value = None)
wb_1.insert(loc = 628,column = '担当_167日目', value = None)
wb_1.insert(loc = 629,column = '担当_168日目', value = None)
wb_1.insert(loc = 630,column = '担当_169日目', value = None)
wb_1.insert(loc = 631,column = '担当_170日目', value = None)
wb_1.insert(loc = 632,column = '担当_171日目', value = None)
wb_1.insert(loc = 633,column = '担当_172日目', value = None)
wb_1.insert(loc = 634,column = '担当_173日目', value = None)
wb_1.insert(loc = 635,column = '担当_174日目', value = None)
wb_1.insert(loc = 636,column = '担当_175日目', value = None)
wb_1.insert(loc = 637,column = '担当_176日目', value = None)
wb_1.insert(loc = 638,column = '担当_177日目', value = None)
wb_1.insert(loc = 639,column = '担当_178日目', value = None)
wb_1.insert(loc = 640,column = '担当_179日目', value = None)
wb_1.insert(loc = 641,column = '担当_180日目', value = None)
wb_1.insert(loc = 642,column = '担当_181日目', value = None)
wb_1.insert(loc = 643,column = '担当_182日目', value = None)
wb_1.insert(loc = 644,column = '担当_183日目', value = None)
wb_1.insert(loc = 645,column = '担当_184日目', value = None)
wb_1.insert(loc = 646,column = '担当_185日目', value = None)
wb_1.insert(loc = 647,column = '担当_186日目', value = None)
wb_1.insert(loc = 648,column = '担当_187日目', value = None)
wb_1.insert(loc = 649,column = '担当_188日目', value = None)
wb_1.insert(loc = 650,column = '担当_189日目', value = None)
wb_1.insert(loc = 651,column = '担当_190日目', value = None)
wb_1.insert(loc = 652,column = '担当_191日目', value = None)
wb_1.insert(loc = 653,column = '担当_192日目', value = None)
wb_1.insert(loc = 654,column = '担当_193日目', value = None)
wb_1.insert(loc = 655,column = '担当_194日目', value = None)
wb_1.insert(loc = 656,column = '担当_195日目', value = None)
wb_1.insert(loc = 657,column = '担当_196日目', value = None)
wb_1.insert(loc = 658,column = '担当_197日目', value = None)
wb_1.insert(loc = 659,column = '担当_198日目', value = None)
wb_1.insert(loc = 660,column = '担当_199日目', value = None)
wb_1.insert(loc = 661,column = '担当_200日目', value = None)
wb_1.insert(loc = 662,column = '担当_201日目', value = None)
wb_1.insert(loc = 663,column = '担当_202日目', value = None)
wb_1.insert(loc = 664,column = '担当_203日目', value = None)
wb_1.insert(loc = 665,column = '担当_204日目', value = None)
wb_1.insert(loc = 666,column = '担当_205日目', value = None)
wb_1.insert(loc = 667,column = '担当_206日目', value = None)
wb_1.insert(loc = 668,column = '担当_207日目', value = None)
wb_1.insert(loc = 669,column = '担当_208日目', value = None)
wb_1.insert(loc = 670,column = '担当_209日目', value = None)
wb_1.insert(loc = 671,column = '担当_210日目', value = None)
wb_1.insert(loc = 672,column = '担当_211日目', value = None)
wb_1.insert(loc = 673,column = '担当_212日目', value = None)

# excelファイルの内容を読み込み
wb_2 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/2フォーマット(渡部さん)/【G】顧客情報(来店時).xlsx', sheet_name='【G】顧客情報', header=1)
# 不足している列を追加
wb_2.insert(loc = 132,column = '日付_131日目', value = None)
wb_2.insert(loc = 133,column = '日付_132日目', value = None)
wb_2.insert(loc = 134,column = '日付_133日目', value = None)
wb_2.insert(loc = 135,column = '日付_134日目', value = None)
wb_2.insert(loc = 136,column = '日付_135日目', value = None)
wb_2.insert(loc = 137,column = '日付_136日目', value = None)
wb_2.insert(loc = 138,column = '日付_137日目', value = None)
wb_2.insert(loc = 139,column = '日付_138日目', value = None)
wb_2.insert(loc = 140,column = '日付_139日目', value = None)
wb_2.insert(loc = 141,column = '日付_140日目', value = None)
wb_2.insert(loc = 142,column = '日付_141日目', value = None)
wb_2.insert(loc = 143,column = '日付_142日目', value = None)
wb_2.insert(loc = 144,column = '日付_143日目', value = None)
wb_2.insert(loc = 145,column = '日付_144日目', value = None)
wb_2.insert(loc = 146,column = '日付_145日目', value = None)
wb_2.insert(loc = 147,column = '日付_146日目', value = None)
wb_2.insert(loc = 148,column = '日付_147日目', value = None)
wb_2.insert(loc = 149,column = '日付_148日目', value = None)
wb_2.insert(loc = 150,column = '日付_149日目', value = None)
wb_2.insert(loc = 151,column = '日付_150日目', value = None)
wb_2.insert(loc = 152,column = '日付_151日目', value = None)
wb_2.insert(loc = 153,column = '日付_152日目', value = None)
wb_2.insert(loc = 154,column = '日付_153日目', value = None)
wb_2.insert(loc = 155,column = '日付_154日目', value = None)
wb_2.insert(loc = 156,column = '日付_155日目', value = None)
wb_2.insert(loc = 157,column = '日付_156日目', value = None)
wb_2.insert(loc = 158,column = '日付_157日目', value = None)
wb_2.insert(loc = 159,column = '日付_158日目', value = None)
wb_2.insert(loc = 160,column = '日付_159日目', value = None)
wb_2.insert(loc = 161,column = '日付_160日目', value = None)
wb_2.insert(loc = 162,column = '日付_161日目', value = None)
wb_2.insert(loc = 163,column = '日付_162日目', value = None)
wb_2.insert(loc = 164,column = '日付_163日目', value = None)
wb_2.insert(loc = 165,column = '日付_164日目', value = None)
wb_2.insert(loc = 166,column = '日付_165日目', value = None)
wb_2.insert(loc = 167,column = '日付_166日目', value = None)
wb_2.insert(loc = 168,column = '日付_167日目', value = None)
wb_2.insert(loc = 169,column = '日付_168日目', value = None)
wb_2.insert(loc = 170,column = '日付_169日目', value = None)
wb_2.insert(loc = 171,column = '日付_170日目', value = None)
wb_2.insert(loc = 172,column = '日付_171日目', value = None)
wb_2.insert(loc = 173,column = '日付_172日目', value = None)
wb_2.insert(loc = 174,column = '日付_173日目', value = None)
wb_2.insert(loc = 175,column = '日付_174日目', value = None)
wb_2.insert(loc = 176,column = '日付_175日目', value = None)
wb_2.insert(loc = 177,column = '日付_176日目', value = None)
wb_2.insert(loc = 178,column = '日付_177日目', value = None)
wb_2.insert(loc = 179,column = '日付_178日目', value = None)
wb_2.insert(loc = 180,column = '日付_179日目', value = None)
wb_2.insert(loc = 181,column = '日付_180日目', value = None)
wb_2.insert(loc = 182,column = '日付_181日目', value = None)
wb_2.insert(loc = 183,column = '日付_182日目', value = None)
wb_2.insert(loc = 184,column = '日付_183日目', value = None)
wb_2.insert(loc = 185,column = '日付_184日目', value = None)
wb_2.insert(loc = 186,column = '日付_185日目', value = None)
wb_2.insert(loc = 187,column = '日付_186日目', value = None)
wb_2.insert(loc = 188,column = '日付_187日目', value = None)
wb_2.insert(loc = 189,column = '日付_188日目', value = None)
wb_2.insert(loc = 190,column = '日付_189日目', value = None)
wb_2.insert(loc = 191,column = '日付_190日目', value = None)
wb_2.insert(loc = 192,column = '日付_191日目', value = None)
wb_2.insert(loc = 193,column = '日付_192日目', value = None)
wb_2.insert(loc = 194,column = '日付_193日目', value = None)
wb_2.insert(loc = 195,column = '日付_194日目', value = None)
wb_2.insert(loc = 196,column = '日付_195日目', value = None)
wb_2.insert(loc = 197,column = '日付_196日目', value = None)
wb_2.insert(loc = 198,column = '日付_197日目', value = None)
wb_2.insert(loc = 199,column = '日付_198日目', value = None)
wb_2.insert(loc = 200,column = '日付_199日目', value = None)
wb_2.insert(loc = 201,column = '日付_200日目', value = None)
wb_2.insert(loc = 202,column = '日付_201日目', value = None)
wb_2.insert(loc = 203,column = '日付_202日目', value = None)
wb_2.insert(loc = 204,column = '日付_203日目', value = None)
wb_2.insert(loc = 205,column = '日付_204日目', value = None)
wb_2.insert(loc = 206,column = '日付_205日目', value = None)
wb_2.insert(loc = 207,column = '日付_206日目', value = None)
wb_2.insert(loc = 208,column = '日付_207日目', value = None)
wb_2.insert(loc = 209,column = '日付_208日目', value = None)
wb_2.insert(loc = 210,column = '日付_209日目', value = None)
wb_2.insert(loc = 211,column = '日付_210日目', value = None)
wb_2.insert(loc = 212,column = '日付_211日目', value = None)
wb_2.insert(loc = 213,column = '日付_212日目', value = None)
wb_2.insert(loc = 214,column = '日付_213日目', value = None)
wb_2.insert(loc = 215,column = '日付_214日目', value = None)
wb_2.insert(loc = 216,column = '日付_215日目', value = None)
wb_2.insert(loc = 217,column = '日付_216日目', value = None)
wb_2.insert(loc = 218,column = '日付_217日目', value = None)
wb_2.insert(loc = 219,column = '日付_218日目', value = None)
wb_2.insert(loc = 220,column = '日付_219日目', value = None)
wb_2.insert(loc = 221,column = '日付_220日目', value = None)
wb_2.insert(loc = 222,column = '日付_221日目', value = None)
wb_2.insert(loc = 223,column = '日付_222日目', value = None)
wb_2.insert(loc = 224,column = '日付_223日目', value = None)
wb_2.insert(loc = 225,column = '日付_224日目', value = None)
wb_2.insert(loc = 226,column = '日付_225日目', value = None)
wb_2.insert(loc = 227,column = '日付_226日目', value = None)
wb_2.insert(loc = 228,column = '日付_227日目', value = None)
wb_2.insert(loc = 229,column = '日付_228日目', value = None)
wb_2.insert(loc = 230,column = '日付_229日目', value = None)
wb_2.insert(loc = 231,column = '日付_230日目', value = None)
wb_2.insert(loc = 362,column = '内容_131日目', value = None)
wb_2.insert(loc = 363,column = '内容_132日目', value = None)
wb_2.insert(loc = 364,column = '内容_133日目', value = None)
wb_2.insert(loc = 365,column = '内容_134日目', value = None)
wb_2.insert(loc = 366,column = '内容_135日目', value = None)
wb_2.insert(loc = 367,column = '内容_136日目', value = None)
wb_2.insert(loc = 368,column = '内容_137日目', value = None)
wb_2.insert(loc = 369,column = '内容_138日目', value = None)
wb_2.insert(loc = 370,column = '内容_139日目', value = None)
wb_2.insert(loc = 371,column = '内容_140日目', value = None)
wb_2.insert(loc = 372,column = '内容_141日目', value = None)
wb_2.insert(loc = 373,column = '内容_142日目', value = None)
wb_2.insert(loc = 374,column = '内容_143日目', value = None)
wb_2.insert(loc = 375,column = '内容_144日目', value = None)
wb_2.insert(loc = 376,column = '内容_145日目', value = None)
wb_2.insert(loc = 377,column = '内容_146日目', value = None)
wb_2.insert(loc = 378,column = '内容_147日目', value = None)
wb_2.insert(loc = 379,column = '内容_148日目', value = None)
wb_2.insert(loc = 380,column = '内容_149日目', value = None)
wb_2.insert(loc = 381,column = '内容_150日目', value = None)
wb_2.insert(loc = 382,column = '内容_151日目', value = None)
wb_2.insert(loc = 383,column = '内容_152日目', value = None)
wb_2.insert(loc = 384,column = '内容_153日目', value = None)
wb_2.insert(loc = 385,column = '内容_154日目', value = None)
wb_2.insert(loc = 386,column = '内容_155日目', value = None)
wb_2.insert(loc = 387,column = '内容_156日目', value = None)
wb_2.insert(loc = 388,column = '内容_157日目', value = None)
wb_2.insert(loc = 389,column = '内容_158日目', value = None)
wb_2.insert(loc = 390,column = '内容_159日目', value = None)
wb_2.insert(loc = 391,column = '内容_160日目', value = None)
wb_2.insert(loc = 392,column = '内容_161日目', value = None)
wb_2.insert(loc = 393,column = '内容_162日目', value = None)
wb_2.insert(loc = 394,column = '内容_163日目', value = None)
wb_2.insert(loc = 395,column = '内容_164日目', value = None)
wb_2.insert(loc = 396,column = '内容_165日目', value = None)
wb_2.insert(loc = 397,column = '内容_166日目', value = None)
wb_2.insert(loc = 398,column = '内容_167日目', value = None)
wb_2.insert(loc = 399,column = '内容_168日目', value = None)
wb_2.insert(loc = 400,column = '内容_169日目', value = None)
wb_2.insert(loc = 401,column = '内容_170日目', value = None)
wb_2.insert(loc = 402,column = '内容_171日目', value = None)
wb_2.insert(loc = 403,column = '内容_172日目', value = None)
wb_2.insert(loc = 404,column = '内容_173日目', value = None)
wb_2.insert(loc = 405,column = '内容_174日目', value = None)
wb_2.insert(loc = 406,column = '内容_175日目', value = None)
wb_2.insert(loc = 407,column = '内容_176日目', value = None)
wb_2.insert(loc = 408,column = '内容_177日目', value = None)
wb_2.insert(loc = 409,column = '内容_178日目', value = None)
wb_2.insert(loc = 410,column = '内容_179日目', value = None)
wb_2.insert(loc = 411,column = '内容_180日目', value = None)
wb_2.insert(loc = 412,column = '内容_181日目', value = None)
wb_2.insert(loc = 413,column = '内容_182日目', value = None)
wb_2.insert(loc = 414,column = '内容_183日目', value = None)
wb_2.insert(loc = 415,column = '内容_184日目', value = None)
wb_2.insert(loc = 416,column = '内容_185日目', value = None)
wb_2.insert(loc = 417,column = '内容_186日目', value = None)
wb_2.insert(loc = 418,column = '内容_187日目', value = None)
wb_2.insert(loc = 419,column = '内容_188日目', value = None)
wb_2.insert(loc = 420,column = '内容_189日目', value = None)
wb_2.insert(loc = 421,column = '内容_190日目', value = None)
wb_2.insert(loc = 422,column = '内容_191日目', value = None)
wb_2.insert(loc = 423,column = '内容_192日目', value = None)
wb_2.insert(loc = 424,column = '内容_193日目', value = None)
wb_2.insert(loc = 425,column = '内容_194日目', value = None)
wb_2.insert(loc = 426,column = '内容_195日目', value = None)
wb_2.insert(loc = 427,column = '内容_196日目', value = None)
wb_2.insert(loc = 428,column = '内容_197日目', value = None)
wb_2.insert(loc = 429,column = '内容_198日目', value = None)
wb_2.insert(loc = 430,column = '内容_199日目', value = None)
wb_2.insert(loc = 431,column = '内容_200日目', value = None)
wb_2.insert(loc = 432,column = '内容_201日目', value = None)
wb_2.insert(loc = 433,column = '内容_202日目', value = None)
wb_2.insert(loc = 434,column = '内容_203日目', value = None)
wb_2.insert(loc = 435,column = '内容_204日目', value = None)
wb_2.insert(loc = 436,column = '内容_205日目', value = None)
wb_2.insert(loc = 437,column = '内容_206日目', value = None)
wb_2.insert(loc = 438,column = '内容_207日目', value = None)
wb_2.insert(loc = 439,column = '内容_208日目', value = None)
wb_2.insert(loc = 440,column = '内容_209日目', value = None)
wb_2.insert(loc = 441,column = '内容_210日目', value = None)
wb_2.insert(loc = 442,column = '内容_211日目', value = None)
wb_2.insert(loc = 443,column = '内容_212日目', value = None)
wb_2.insert(loc = 444,column = '内容_213日目', value = None)
wb_2.insert(loc = 445,column = '内容_214日目', value = None)
wb_2.insert(loc = 446,column = '内容_215日目', value = None)
wb_2.insert(loc = 447,column = '内容_216日目', value = None)
wb_2.insert(loc = 448,column = '内容_217日目', value = None)
wb_2.insert(loc = 449,column = '内容_218日目', value = None)
wb_2.insert(loc = 450,column = '内容_219日目', value = None)
wb_2.insert(loc = 451,column = '内容_220日目', value = None)
wb_2.insert(loc = 452,column = '内容_221日目', value = None)
wb_2.insert(loc = 453,column = '内容_222日目', value = None)
wb_2.insert(loc = 454,column = '内容_223日目', value = None)
wb_2.insert(loc = 455,column = '内容_224日目', value = None)
wb_2.insert(loc = 456,column = '内容_225日目', value = None)
wb_2.insert(loc = 457,column = '内容_226日目', value = None)
wb_2.insert(loc = 458,column = '内容_227日目', value = None)
wb_2.insert(loc = 459,column = '内容_228日目', value = None)
wb_2.insert(loc = 460,column = '内容_229日目', value = None)
wb_2.insert(loc = 461,column = '内容_230日目', value = None)
wb_2.insert(loc = 592,column = '担当_131日目', value = None)
wb_2.insert(loc = 593,column = '担当_132日目', value = None)
wb_2.insert(loc = 594,column = '担当_133日目', value = None)
wb_2.insert(loc = 595,column = '担当_134日目', value = None)
wb_2.insert(loc = 596,column = '担当_135日目', value = None)
wb_2.insert(loc = 597,column = '担当_136日目', value = None)
wb_2.insert(loc = 598,column = '担当_137日目', value = None)
wb_2.insert(loc = 599,column = '担当_138日目', value = None)
wb_2.insert(loc = 600,column = '担当_139日目', value = None)
wb_2.insert(loc = 601,column = '担当_140日目', value = None)
wb_2.insert(loc = 602,column = '担当_141日目', value = None)
wb_2.insert(loc = 603,column = '担当_142日目', value = None)
wb_2.insert(loc = 604,column = '担当_143日目', value = None)
wb_2.insert(loc = 605,column = '担当_144日目', value = None)
wb_2.insert(loc = 606,column = '担当_145日目', value = None)
wb_2.insert(loc = 607,column = '担当_146日目', value = None)
wb_2.insert(loc = 608,column = '担当_147日目', value = None)
wb_2.insert(loc = 609,column = '担当_148日目', value = None)
wb_2.insert(loc = 610,column = '担当_149日目', value = None)
wb_2.insert(loc = 611,column = '担当_150日目', value = None)
wb_2.insert(loc = 612,column = '担当_151日目', value = None)
wb_2.insert(loc = 613,column = '担当_152日目', value = None)
wb_2.insert(loc = 614,column = '担当_153日目', value = None)
wb_2.insert(loc = 615,column = '担当_154日目', value = None)
wb_2.insert(loc = 616,column = '担当_155日目', value = None)
wb_2.insert(loc = 617,column = '担当_156日目', value = None)
wb_2.insert(loc = 618,column = '担当_157日目', value = None)
wb_2.insert(loc = 619,column = '担当_158日目', value = None)
wb_2.insert(loc = 620,column = '担当_159日目', value = None)
wb_2.insert(loc = 621,column = '担当_160日目', value = None)
wb_2.insert(loc = 622,column = '担当_161日目', value = None)
wb_2.insert(loc = 623,column = '担当_162日目', value = None)
wb_2.insert(loc = 624,column = '担当_163日目', value = None)
wb_2.insert(loc = 625,column = '担当_164日目', value = None)
wb_2.insert(loc = 626,column = '担当_165日目', value = None)
wb_2.insert(loc = 627,column = '担当_166日目', value = None)
wb_2.insert(loc = 628,column = '担当_167日目', value = None)
wb_2.insert(loc = 629,column = '担当_168日目', value = None)
wb_2.insert(loc = 630,column = '担当_169日目', value = None)
wb_2.insert(loc = 631,column = '担当_170日目', value = None)
wb_2.insert(loc = 632,column = '担当_171日目', value = None)
wb_2.insert(loc = 633,column = '担当_172日目', value = None)
wb_2.insert(loc = 634,column = '担当_173日目', value = None)
wb_2.insert(loc = 635,column = '担当_174日目', value = None)
wb_2.insert(loc = 636,column = '担当_175日目', value = None)
wb_2.insert(loc = 637,column = '担当_176日目', value = None)
wb_2.insert(loc = 638,column = '担当_177日目', value = None)
wb_2.insert(loc = 639,column = '担当_178日目', value = None)
wb_2.insert(loc = 640,column = '担当_179日目', value = None)
wb_2.insert(loc = 641,column = '担当_180日目', value = None)
wb_2.insert(loc = 642,column = '担当_181日目', value = None)
wb_2.insert(loc = 643,column = '担当_182日目', value = None)
wb_2.insert(loc = 644,column = '担当_183日目', value = None)
wb_2.insert(loc = 645,column = '担当_184日目', value = None)
wb_2.insert(loc = 646,column = '担当_185日目', value = None)
wb_2.insert(loc = 647,column = '担当_186日目', value = None)
wb_2.insert(loc = 648,column = '担当_187日目', value = None)
wb_2.insert(loc = 649,column = '担当_188日目', value = None)
wb_2.insert(loc = 650,column = '担当_189日目', value = None)
wb_2.insert(loc = 651,column = '担当_190日目', value = None)
wb_2.insert(loc = 652,column = '担当_191日目', value = None)
wb_2.insert(loc = 653,column = '担当_192日目', value = None)
wb_2.insert(loc = 654,column = '担当_193日目', value = None)
wb_2.insert(loc = 655,column = '担当_194日目', value = None)
wb_2.insert(loc = 656,column = '担当_195日目', value = None)
wb_2.insert(loc = 657,column = '担当_196日目', value = None)
wb_2.insert(loc = 658,column = '担当_197日目', value = None)
wb_2.insert(loc = 659,column = '担当_198日目', value = None)
wb_2.insert(loc = 660,column = '担当_199日目', value = None)
wb_2.insert(loc = 661,column = '担当_200日目', value = None)
wb_2.insert(loc = 662,column = '担当_201日目', value = None)
wb_2.insert(loc = 663,column = '担当_202日目', value = None)
wb_2.insert(loc = 664,column = '担当_203日目', value = None)
wb_2.insert(loc = 665,column = '担当_204日目', value = None)
wb_2.insert(loc = 666,column = '担当_205日目', value = None)
wb_2.insert(loc = 667,column = '担当_206日目', value = None)
wb_2.insert(loc = 668,column = '担当_207日目', value = None)
wb_2.insert(loc = 669,column = '担当_208日目', value = None)
wb_2.insert(loc = 670,column = '担当_209日目', value = None)
wb_2.insert(loc = 671,column = '担当_210日目', value = None)
wb_2.insert(loc = 672,column = '担当_211日目', value = None)
wb_2.insert(loc = 673,column = '担当_212日目', value = None)

# excelファイルの内容を読み込み
wb_3 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/2フォーマット(柏葉さん)/【G】顧客情報(来店時).xlsx', sheet_name='【G】顧客情報', header=1)
wb_4 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/慶野さん/【G】顧客情報(来店時).xlsx', sheet_name='【G】顧客情報', header=1)
# 不足している列を追加
wb_4.insert(loc = 118,column = '日付_117日目', value = None)
wb_4.insert(loc = 119,column = '日付_118日目', value = None)
wb_4.insert(loc = 120,column = '日付_119日目', value = None)
wb_4.insert(loc = 121,column = '日付_120日目', value = None)
wb_4.insert(loc = 122,column = '日付_121日目', value = None)
wb_4.insert(loc = 123,column = '日付_122日目', value = None)
wb_4.insert(loc = 124,column = '日付_123日目', value = None)
wb_4.insert(loc = 125,column = '日付_124日目', value = None)
wb_4.insert(loc = 126,column = '日付_125日目', value = None)
wb_4.insert(loc = 127,column = '日付_126日目', value = None)
wb_4.insert(loc = 128,column = '日付_127日目', value = None)
wb_4.insert(loc = 129,column = '日付_128日目', value = None)
wb_4.insert(loc = 130,column = '日付_129日目', value = None)
wb_4.insert(loc = 131,column = '日付_130日目', value = None)
wb_4.insert(loc = 132,column = '日付_131日目', value = None)
wb_4.insert(loc = 133,column = '日付_132日目', value = None)
wb_4.insert(loc = 134,column = '日付_133日目', value = None)
wb_4.insert(loc = 135,column = '日付_134日目', value = None)
wb_4.insert(loc = 136,column = '日付_135日目', value = None)
wb_4.insert(loc = 137,column = '日付_136日目', value = None)
wb_4.insert(loc = 138,column = '日付_137日目', value = None)
wb_4.insert(loc = 139,column = '日付_138日目', value = None)
wb_4.insert(loc = 140,column = '日付_139日目', value = None)
wb_4.insert(loc = 141,column = '日付_140日目', value = None)
wb_4.insert(loc = 142,column = '日付_141日目', value = None)
wb_4.insert(loc = 143,column = '日付_142日目', value = None)
wb_4.insert(loc = 144,column = '日付_143日目', value = None)
wb_4.insert(loc = 145,column = '日付_144日目', value = None)
wb_4.insert(loc = 146,column = '日付_145日目', value = None)
wb_4.insert(loc = 147,column = '日付_146日目', value = None)
wb_4.insert(loc = 148,column = '日付_147日目', value = None)
wb_4.insert(loc = 149,column = '日付_148日目', value = None)
wb_4.insert(loc = 150,column = '日付_149日目', value = None)
wb_4.insert(loc = 151,column = '日付_150日目', value = None)
wb_4.insert(loc = 152,column = '日付_151日目', value = None)
wb_4.insert(loc = 153,column = '日付_152日目', value = None)
wb_4.insert(loc = 154,column = '日付_153日目', value = None)
wb_4.insert(loc = 155,column = '日付_154日目', value = None)
wb_4.insert(loc = 156,column = '日付_155日目', value = None)
wb_4.insert(loc = 157,column = '日付_156日目', value = None)
wb_4.insert(loc = 158,column = '日付_157日目', value = None)
wb_4.insert(loc = 159,column = '日付_158日目', value = None)
wb_4.insert(loc = 160,column = '日付_159日目', value = None)
wb_4.insert(loc = 161,column = '日付_160日目', value = None)
wb_4.insert(loc = 162,column = '日付_161日目', value = None)
wb_4.insert(loc = 163,column = '日付_162日目', value = None)
wb_4.insert(loc = 164,column = '日付_163日目', value = None)
wb_4.insert(loc = 165,column = '日付_164日目', value = None)
wb_4.insert(loc = 166,column = '日付_165日目', value = None)
wb_4.insert(loc = 167,column = '日付_166日目', value = None)
wb_4.insert(loc = 168,column = '日付_167日目', value = None)
wb_4.insert(loc = 169,column = '日付_168日目', value = None)
wb_4.insert(loc = 170,column = '日付_169日目', value = None)
wb_4.insert(loc = 171,column = '日付_170日目', value = None)
wb_4.insert(loc = 172,column = '日付_171日目', value = None)
wb_4.insert(loc = 173,column = '日付_172日目', value = None)
wb_4.insert(loc = 174,column = '日付_173日目', value = None)
wb_4.insert(loc = 175,column = '日付_174日目', value = None)
wb_4.insert(loc = 176,column = '日付_175日目', value = None)
wb_4.insert(loc = 177,column = '日付_176日目', value = None)
wb_4.insert(loc = 178,column = '日付_177日目', value = None)
wb_4.insert(loc = 179,column = '日付_178日目', value = None)
wb_4.insert(loc = 180,column = '日付_179日目', value = None)
wb_4.insert(loc = 181,column = '日付_180日目', value = None)
wb_4.insert(loc = 182,column = '日付_181日目', value = None)
wb_4.insert(loc = 183,column = '日付_182日目', value = None)
wb_4.insert(loc = 184,column = '日付_183日目', value = None)
wb_4.insert(loc = 185,column = '日付_184日目', value = None)
wb_4.insert(loc = 186,column = '日付_185日目', value = None)
wb_4.insert(loc = 187,column = '日付_186日目', value = None)
wb_4.insert(loc = 188,column = '日付_187日目', value = None)
wb_4.insert(loc = 189,column = '日付_188日目', value = None)
wb_4.insert(loc = 190,column = '日付_189日目', value = None)
wb_4.insert(loc = 191,column = '日付_190日目', value = None)
wb_4.insert(loc = 192,column = '日付_191日目', value = None)
wb_4.insert(loc = 193,column = '日付_192日目', value = None)
wb_4.insert(loc = 194,column = '日付_193日目', value = None)
wb_4.insert(loc = 195,column = '日付_194日目', value = None)
wb_4.insert(loc = 196,column = '日付_195日目', value = None)
wb_4.insert(loc = 197,column = '日付_196日目', value = None)
wb_4.insert(loc = 198,column = '日付_197日目', value = None)
wb_4.insert(loc = 199,column = '日付_198日目', value = None)
wb_4.insert(loc = 200,column = '日付_199日目', value = None)
wb_4.insert(loc = 201,column = '日付_200日目', value = None)
wb_4.insert(loc = 202,column = '日付_201日目', value = None)
wb_4.insert(loc = 203,column = '日付_202日目', value = None)
wb_4.insert(loc = 204,column = '日付_203日目', value = None)
wb_4.insert(loc = 205,column = '日付_204日目', value = None)
wb_4.insert(loc = 206,column = '日付_205日目', value = None)
wb_4.insert(loc = 207,column = '日付_206日目', value = None)
wb_4.insert(loc = 208,column = '日付_207日目', value = None)
wb_4.insert(loc = 209,column = '日付_208日目', value = None)
wb_4.insert(loc = 210,column = '日付_209日目', value = None)
wb_4.insert(loc = 211,column = '日付_210日目', value = None)
wb_4.insert(loc = 212,column = '日付_211日目', value = None)
wb_4.insert(loc = 213,column = '日付_212日目', value = None)
wb_4.insert(loc = 214,column = '日付_213日目', value = None)
wb_4.insert(loc = 215,column = '日付_214日目', value = None)
wb_4.insert(loc = 216,column = '日付_215日目', value = None)
wb_4.insert(loc = 217,column = '日付_216日目', value = None)
wb_4.insert(loc = 218,column = '日付_217日目', value = None)
wb_4.insert(loc = 219,column = '日付_218日目', value = None)
wb_4.insert(loc = 220,column = '日付_219日目', value = None)
wb_4.insert(loc = 221,column = '日付_220日目', value = None)
wb_4.insert(loc = 222,column = '日付_221日目', value = None)
wb_4.insert(loc = 223,column = '日付_222日目', value = None)
wb_4.insert(loc = 224,column = '日付_223日目', value = None)
wb_4.insert(loc = 225,column = '日付_224日目', value = None)
wb_4.insert(loc = 226,column = '日付_225日目', value = None)
wb_4.insert(loc = 227,column = '日付_226日目', value = None)
wb_4.insert(loc = 228,column = '日付_227日目', value = None)
wb_4.insert(loc = 229,column = '日付_228日目', value = None)
wb_4.insert(loc = 230,column = '日付_229日目', value = None)
wb_4.insert(loc = 231,column = '日付_230日目', value = None)
wb_4.insert(loc = 348,column = '内容_117日目', value = None)
wb_4.insert(loc = 349,column = '内容_118日目', value = None)
wb_4.insert(loc = 350,column = '内容_119日目', value = None)
wb_4.insert(loc = 351,column = '内容_120日目', value = None)
wb_4.insert(loc = 352,column = '内容_121日目', value = None)
wb_4.insert(loc = 353,column = '内容_122日目', value = None)
wb_4.insert(loc = 354,column = '内容_123日目', value = None)
wb_4.insert(loc = 355,column = '内容_124日目', value = None)
wb_4.insert(loc = 356,column = '内容_125日目', value = None)
wb_4.insert(loc = 357,column = '内容_126日目', value = None)
wb_4.insert(loc = 358,column = '内容_127日目', value = None)
wb_4.insert(loc = 359,column = '内容_128日目', value = None)
wb_4.insert(loc = 360,column = '内容_129日目', value = None)
wb_4.insert(loc = 361,column = '内容_130日目', value = None)
wb_4.insert(loc = 362,column = '内容_131日目', value = None)
wb_4.insert(loc = 363,column = '内容_132日目', value = None)
wb_4.insert(loc = 364,column = '内容_133日目', value = None)
wb_4.insert(loc = 365,column = '内容_134日目', value = None)
wb_4.insert(loc = 366,column = '内容_135日目', value = None)
wb_4.insert(loc = 367,column = '内容_136日目', value = None)
wb_4.insert(loc = 368,column = '内容_137日目', value = None)
wb_4.insert(loc = 369,column = '内容_138日目', value = None)
wb_4.insert(loc = 370,column = '内容_139日目', value = None)
wb_4.insert(loc = 371,column = '内容_140日目', value = None)
wb_4.insert(loc = 372,column = '内容_141日目', value = None)
wb_4.insert(loc = 373,column = '内容_142日目', value = None)
wb_4.insert(loc = 374,column = '内容_143日目', value = None)
wb_4.insert(loc = 375,column = '内容_144日目', value = None)
wb_4.insert(loc = 376,column = '内容_145日目', value = None)
wb_4.insert(loc = 377,column = '内容_146日目', value = None)
wb_4.insert(loc = 378,column = '内容_147日目', value = None)
wb_4.insert(loc = 379,column = '内容_148日目', value = None)
wb_4.insert(loc = 380,column = '内容_149日目', value = None)
wb_4.insert(loc = 381,column = '内容_150日目', value = None)
wb_4.insert(loc = 382,column = '内容_151日目', value = None)
wb_4.insert(loc = 383,column = '内容_152日目', value = None)
wb_4.insert(loc = 384,column = '内容_153日目', value = None)
wb_4.insert(loc = 385,column = '内容_154日目', value = None)
wb_4.insert(loc = 386,column = '内容_155日目', value = None)
wb_4.insert(loc = 387,column = '内容_156日目', value = None)
wb_4.insert(loc = 388,column = '内容_157日目', value = None)
wb_4.insert(loc = 389,column = '内容_158日目', value = None)
wb_4.insert(loc = 390,column = '内容_159日目', value = None)
wb_4.insert(loc = 391,column = '内容_160日目', value = None)
wb_4.insert(loc = 392,column = '内容_161日目', value = None)
wb_4.insert(loc = 393,column = '内容_162日目', value = None)
wb_4.insert(loc = 394,column = '内容_163日目', value = None)
wb_4.insert(loc = 395,column = '内容_164日目', value = None)
wb_4.insert(loc = 396,column = '内容_165日目', value = None)
wb_4.insert(loc = 397,column = '内容_166日目', value = None)
wb_4.insert(loc = 398,column = '内容_167日目', value = None)
wb_4.insert(loc = 399,column = '内容_168日目', value = None)
wb_4.insert(loc = 400,column = '内容_169日目', value = None)
wb_4.insert(loc = 401,column = '内容_170日目', value = None)
wb_4.insert(loc = 402,column = '内容_171日目', value = None)
wb_4.insert(loc = 403,column = '内容_172日目', value = None)
wb_4.insert(loc = 404,column = '内容_173日目', value = None)
wb_4.insert(loc = 405,column = '内容_174日目', value = None)
wb_4.insert(loc = 406,column = '内容_175日目', value = None)
wb_4.insert(loc = 407,column = '内容_176日目', value = None)
wb_4.insert(loc = 408,column = '内容_177日目', value = None)
wb_4.insert(loc = 409,column = '内容_178日目', value = None)
wb_4.insert(loc = 410,column = '内容_179日目', value = None)
wb_4.insert(loc = 411,column = '内容_180日目', value = None)
wb_4.insert(loc = 412,column = '内容_181日目', value = None)
wb_4.insert(loc = 413,column = '内容_182日目', value = None)
wb_4.insert(loc = 414,column = '内容_183日目', value = None)
wb_4.insert(loc = 415,column = '内容_184日目', value = None)
wb_4.insert(loc = 416,column = '内容_185日目', value = None)
wb_4.insert(loc = 417,column = '内容_186日目', value = None)
wb_4.insert(loc = 418,column = '内容_187日目', value = None)
wb_4.insert(loc = 419,column = '内容_188日目', value = None)
wb_4.insert(loc = 420,column = '内容_189日目', value = None)
wb_4.insert(loc = 421,column = '内容_190日目', value = None)
wb_4.insert(loc = 422,column = '内容_191日目', value = None)
wb_4.insert(loc = 423,column = '内容_192日目', value = None)
wb_4.insert(loc = 424,column = '内容_193日目', value = None)
wb_4.insert(loc = 425,column = '内容_194日目', value = None)
wb_4.insert(loc = 426,column = '内容_195日目', value = None)
wb_4.insert(loc = 427,column = '内容_196日目', value = None)
wb_4.insert(loc = 428,column = '内容_197日目', value = None)
wb_4.insert(loc = 429,column = '内容_198日目', value = None)
wb_4.insert(loc = 430,column = '内容_199日目', value = None)
wb_4.insert(loc = 431,column = '内容_200日目', value = None)
wb_4.insert(loc = 432,column = '内容_201日目', value = None)
wb_4.insert(loc = 433,column = '内容_202日目', value = None)
wb_4.insert(loc = 434,column = '内容_203日目', value = None)
wb_4.insert(loc = 435,column = '内容_204日目', value = None)
wb_4.insert(loc = 436,column = '内容_205日目', value = None)
wb_4.insert(loc = 437,column = '内容_206日目', value = None)
wb_4.insert(loc = 438,column = '内容_207日目', value = None)
wb_4.insert(loc = 439,column = '内容_208日目', value = None)
wb_4.insert(loc = 440,column = '内容_209日目', value = None)
wb_4.insert(loc = 441,column = '内容_210日目', value = None)
wb_4.insert(loc = 442,column = '内容_211日目', value = None)
wb_4.insert(loc = 443,column = '内容_212日目', value = None)
wb_4.insert(loc = 444,column = '内容_213日目', value = None)
wb_4.insert(loc = 445,column = '内容_214日目', value = None)
wb_4.insert(loc = 446,column = '内容_215日目', value = None)
wb_4.insert(loc = 447,column = '内容_216日目', value = None)
wb_4.insert(loc = 448,column = '内容_217日目', value = None)
wb_4.insert(loc = 449,column = '内容_218日目', value = None)
wb_4.insert(loc = 450,column = '内容_219日目', value = None)
wb_4.insert(loc = 451,column = '内容_220日目', value = None)
wb_4.insert(loc = 452,column = '内容_221日目', value = None)
wb_4.insert(loc = 453,column = '内容_222日目', value = None)
wb_4.insert(loc = 454,column = '内容_223日目', value = None)
wb_4.insert(loc = 455,column = '内容_224日目', value = None)
wb_4.insert(loc = 456,column = '内容_225日目', value = None)
wb_4.insert(loc = 457,column = '内容_226日目', value = None)
wb_4.insert(loc = 458,column = '内容_227日目', value = None)
wb_4.insert(loc = 459,column = '内容_228日目', value = None)
wb_4.insert(loc = 460,column = '内容_229日目', value = None)
wb_4.insert(loc = 461,column = '内容_230日目', value = None)
wb_4.insert(loc = 578,column = '担当_117日目', value = None)
wb_4.insert(loc = 579,column = '担当_118日目', value = None)
wb_4.insert(loc = 580,column = '担当_119日目', value = None)
wb_4.insert(loc = 581,column = '担当_120日目', value = None)
wb_4.insert(loc = 582,column = '担当_121日目', value = None)
wb_4.insert(loc = 583,column = '担当_122日目', value = None)
wb_4.insert(loc = 584,column = '担当_123日目', value = None)
wb_4.insert(loc = 585,column = '担当_124日目', value = None)
wb_4.insert(loc = 586,column = '担当_125日目', value = None)
wb_4.insert(loc = 587,column = '担当_126日目', value = None)
wb_4.insert(loc = 588,column = '担当_127日目', value = None)
wb_4.insert(loc = 589,column = '担当_128日目', value = None)
wb_4.insert(loc = 590,column = '担当_129日目', value = None)
wb_4.insert(loc = 591,column = '担当_130日目', value = None)
wb_4.insert(loc = 592,column = '担当_131日目', value = None)
wb_4.insert(loc = 593,column = '担当_132日目', value = None)
wb_4.insert(loc = 594,column = '担当_133日目', value = None)
wb_4.insert(loc = 595,column = '担当_134日目', value = None)
wb_4.insert(loc = 596,column = '担当_135日目', value = None)
wb_4.insert(loc = 597,column = '担当_136日目', value = None)
wb_4.insert(loc = 598,column = '担当_137日目', value = None)
wb_4.insert(loc = 599,column = '担当_138日目', value = None)
wb_4.insert(loc = 600,column = '担当_139日目', value = None)
wb_4.insert(loc = 601,column = '担当_140日目', value = None)
wb_4.insert(loc = 602,column = '担当_141日目', value = None)
wb_4.insert(loc = 603,column = '担当_142日目', value = None)
wb_4.insert(loc = 604,column = '担当_143日目', value = None)
wb_4.insert(loc = 605,column = '担当_144日目', value = None)
wb_4.insert(loc = 606,column = '担当_145日目', value = None)
wb_4.insert(loc = 607,column = '担当_146日目', value = None)
wb_4.insert(loc = 608,column = '担当_147日目', value = None)
wb_4.insert(loc = 609,column = '担当_148日目', value = None)
wb_4.insert(loc = 610,column = '担当_149日目', value = None)
wb_4.insert(loc = 611,column = '担当_150日目', value = None)
wb_4.insert(loc = 612,column = '担当_151日目', value = None)
wb_4.insert(loc = 613,column = '担当_152日目', value = None)
wb_4.insert(loc = 614,column = '担当_153日目', value = None)
wb_4.insert(loc = 615,column = '担当_154日目', value = None)
wb_4.insert(loc = 616,column = '担当_155日目', value = None)
wb_4.insert(loc = 617,column = '担当_156日目', value = None)
wb_4.insert(loc = 618,column = '担当_157日目', value = None)
wb_4.insert(loc = 619,column = '担当_158日目', value = None)
wb_4.insert(loc = 620,column = '担当_159日目', value = None)
wb_4.insert(loc = 621,column = '担当_160日目', value = None)
wb_4.insert(loc = 622,column = '担当_161日目', value = None)
wb_4.insert(loc = 623,column = '担当_162日目', value = None)
wb_4.insert(loc = 624,column = '担当_163日目', value = None)
wb_4.insert(loc = 625,column = '担当_164日目', value = None)
wb_4.insert(loc = 626,column = '担当_165日目', value = None)
wb_4.insert(loc = 627,column = '担当_166日目', value = None)
wb_4.insert(loc = 628,column = '担当_167日目', value = None)
wb_4.insert(loc = 629,column = '担当_168日目', value = None)
wb_4.insert(loc = 630,column = '担当_169日目', value = None)
wb_4.insert(loc = 631,column = '担当_170日目', value = None)
wb_4.insert(loc = 632,column = '担当_171日目', value = None)
wb_4.insert(loc = 633,column = '担当_172日目', value = None)
wb_4.insert(loc = 634,column = '担当_173日目', value = None)
wb_4.insert(loc = 635,column = '担当_174日目', value = None)
wb_4.insert(loc = 636,column = '担当_175日目', value = None)
wb_4.insert(loc = 637,column = '担当_176日目', value = None)
wb_4.insert(loc = 638,column = '担当_177日目', value = None)
wb_4.insert(loc = 639,column = '担当_178日目', value = None)
wb_4.insert(loc = 640,column = '担当_179日目', value = None)
wb_4.insert(loc = 641,column = '担当_180日目', value = None)
wb_4.insert(loc = 642,column = '担当_181日目', value = None)
wb_4.insert(loc = 643,column = '担当_182日目', value = None)
wb_4.insert(loc = 644,column = '担当_183日目', value = None)
wb_4.insert(loc = 645,column = '担当_184日目', value = None)
wb_4.insert(loc = 646,column = '担当_185日目', value = None)
wb_4.insert(loc = 647,column = '担当_186日目', value = None)
wb_4.insert(loc = 648,column = '担当_187日目', value = None)
wb_4.insert(loc = 649,column = '担当_188日目', value = None)
wb_4.insert(loc = 650,column = '担当_189日目', value = None)
wb_4.insert(loc = 651,column = '担当_190日目', value = None)
wb_4.insert(loc = 652,column = '担当_191日目', value = None)
wb_4.insert(loc = 653,column = '担当_192日目', value = None)
wb_4.insert(loc = 654,column = '担当_193日目', value = None)
wb_4.insert(loc = 655,column = '担当_194日目', value = None)
wb_4.insert(loc = 656,column = '担当_195日目', value = None)
wb_4.insert(loc = 657,column = '担当_196日目', value = None)
wb_4.insert(loc = 658,column = '担当_197日目', value = None)
wb_4.insert(loc = 659,column = '担当_198日目', value = None)
wb_4.insert(loc = 660,column = '担当_199日目', value = None)
wb_4.insert(loc = 661,column = '担当_200日目', value = None)
wb_4.insert(loc = 662,column = '担当_201日目', value = None)
wb_4.insert(loc = 663,column = '担当_202日目', value = None)
wb_4.insert(loc = 664,column = '担当_203日目', value = None)
wb_4.insert(loc = 665,column = '担当_204日目', value = None)
wb_4.insert(loc = 666,column = '担当_205日目', value = None)
wb_4.insert(loc = 667,column = '担当_206日目', value = None)
wb_4.insert(loc = 668,column = '担当_207日目', value = None)
wb_4.insert(loc = 669,column = '担当_208日目', value = None)
wb_4.insert(loc = 670,column = '担当_209日目', value = None)
wb_4.insert(loc = 671,column = '担当_210日目', value = None)
wb_4.insert(loc = 672,column = '担当_211日目', value = None)
wb_4.insert(loc = 673,column = '担当_212日目', value = None)

# excelファイルの内容を読み込み
wb_5 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/佐藤さん/【G】顧客情報(来店時).xlsx', sheet_name='【G】顧客情報', header=1)
# 不足している列を追加
wb_5.insert(loc = 74,column = '日付_73日目', value = None)
wb_5.insert(loc = 75,column = '日付_74日目', value = None)
wb_5.insert(loc = 76,column = '日付_75日目', value = None)
wb_5.insert(loc = 77,column = '日付_76日目', value = None)
wb_5.insert(loc = 78,column = '日付_77日目', value = None)
wb_5.insert(loc = 79,column = '日付_78日目', value = None)
wb_5.insert(loc = 80,column = '日付_79日目', value = None)
wb_5.insert(loc = 81,column = '日付_80日目', value = None)
wb_5.insert(loc = 82,column = '日付_81日目', value = None)
wb_5.insert(loc = 83,column = '日付_82日目', value = None)
wb_5.insert(loc = 84,column = '日付_83日目', value = None)
wb_5.insert(loc = 85,column = '日付_84日目', value = None)
wb_5.insert(loc = 86,column = '日付_85日目', value = None)
wb_5.insert(loc = 87,column = '日付_86日目', value = None)
wb_5.insert(loc = 88,column = '日付_87日目', value = None)
wb_5.insert(loc = 89,column = '日付_88日目', value = None)
wb_5.insert(loc = 90,column = '日付_89日目', value = None)
wb_5.insert(loc = 91,column = '日付_90日目', value = None)
wb_5.insert(loc = 92,column = '日付_91日目', value = None)
wb_5.insert(loc = 93,column = '日付_92日目', value = None)
wb_5.insert(loc = 94,column = '日付_93日目', value = None)
wb_5.insert(loc = 95,column = '日付_94日目', value = None)
wb_5.insert(loc = 96,column = '日付_95日目', value = None)
wb_5.insert(loc = 97,column = '日付_96日目', value = None)
wb_5.insert(loc = 98,column = '日付_97日目', value = None)
wb_5.insert(loc = 99,column = '日付_98日目', value = None)
wb_5.insert(loc = 100,column = '日付_99日目', value = None)
wb_5.insert(loc = 101,column = '日付_100日目', value = None)
wb_5.insert(loc = 102,column = '日付_101日目', value = None)
wb_5.insert(loc = 103,column = '日付_102日目', value = None)
wb_5.insert(loc = 104,column = '日付_103日目', value = None)
wb_5.insert(loc = 105,column = '日付_104日目', value = None)
wb_5.insert(loc = 106,column = '日付_105日目', value = None)
wb_5.insert(loc = 107,column = '日付_106日目', value = None)
wb_5.insert(loc = 108,column = '日付_107日目', value = None)
wb_5.insert(loc = 109,column = '日付_108日目', value = None)
wb_5.insert(loc = 110,column = '日付_109日目', value = None)
wb_5.insert(loc = 111,column = '日付_110日目', value = None)
wb_5.insert(loc = 112,column = '日付_111日目', value = None)
wb_5.insert(loc = 113,column = '日付_112日目', value = None)
wb_5.insert(loc = 114,column = '日付_113日目', value = None)
wb_5.insert(loc = 115,column = '日付_114日目', value = None)
wb_5.insert(loc = 116,column = '日付_115日目', value = None)
wb_5.insert(loc = 117,column = '日付_116日目', value = None)
wb_5.insert(loc = 118,column = '日付_117日目', value = None)
wb_5.insert(loc = 119,column = '日付_118日目', value = None)
wb_5.insert(loc = 120,column = '日付_119日目', value = None)
wb_5.insert(loc = 121,column = '日付_120日目', value = None)
wb_5.insert(loc = 122,column = '日付_121日目', value = None)
wb_5.insert(loc = 123,column = '日付_122日目', value = None)
wb_5.insert(loc = 124,column = '日付_123日目', value = None)
wb_5.insert(loc = 125,column = '日付_124日目', value = None)
wb_5.insert(loc = 126,column = '日付_125日目', value = None)
wb_5.insert(loc = 127,column = '日付_126日目', value = None)
wb_5.insert(loc = 128,column = '日付_127日目', value = None)
wb_5.insert(loc = 129,column = '日付_128日目', value = None)
wb_5.insert(loc = 130,column = '日付_129日目', value = None)
wb_5.insert(loc = 131,column = '日付_130日目', value = None)
wb_5.insert(loc = 132,column = '日付_131日目', value = None)
wb_5.insert(loc = 133,column = '日付_132日目', value = None)
wb_5.insert(loc = 134,column = '日付_133日目', value = None)
wb_5.insert(loc = 135,column = '日付_134日目', value = None)
wb_5.insert(loc = 136,column = '日付_135日目', value = None)
wb_5.insert(loc = 137,column = '日付_136日目', value = None)
wb_5.insert(loc = 138,column = '日付_137日目', value = None)
wb_5.insert(loc = 139,column = '日付_138日目', value = None)
wb_5.insert(loc = 140,column = '日付_139日目', value = None)
wb_5.insert(loc = 141,column = '日付_140日目', value = None)
wb_5.insert(loc = 142,column = '日付_141日目', value = None)
wb_5.insert(loc = 143,column = '日付_142日目', value = None)
wb_5.insert(loc = 144,column = '日付_143日目', value = None)
wb_5.insert(loc = 145,column = '日付_144日目', value = None)
wb_5.insert(loc = 146,column = '日付_145日目', value = None)
wb_5.insert(loc = 147,column = '日付_146日目', value = None)
wb_5.insert(loc = 148,column = '日付_147日目', value = None)
wb_5.insert(loc = 149,column = '日付_148日目', value = None)
wb_5.insert(loc = 150,column = '日付_149日目', value = None)
wb_5.insert(loc = 151,column = '日付_150日目', value = None)
wb_5.insert(loc = 152,column = '日付_151日目', value = None)
wb_5.insert(loc = 153,column = '日付_152日目', value = None)
wb_5.insert(loc = 154,column = '日付_153日目', value = None)
wb_5.insert(loc = 155,column = '日付_154日目', value = None)
wb_5.insert(loc = 156,column = '日付_155日目', value = None)
wb_5.insert(loc = 157,column = '日付_156日目', value = None)
wb_5.insert(loc = 158,column = '日付_157日目', value = None)
wb_5.insert(loc = 159,column = '日付_158日目', value = None)
wb_5.insert(loc = 160,column = '日付_159日目', value = None)
wb_5.insert(loc = 161,column = '日付_160日目', value = None)
wb_5.insert(loc = 162,column = '日付_161日目', value = None)
wb_5.insert(loc = 163,column = '日付_162日目', value = None)
wb_5.insert(loc = 164,column = '日付_163日目', value = None)
wb_5.insert(loc = 165,column = '日付_164日目', value = None)
wb_5.insert(loc = 166,column = '日付_165日目', value = None)
wb_5.insert(loc = 167,column = '日付_166日目', value = None)
wb_5.insert(loc = 168,column = '日付_167日目', value = None)
wb_5.insert(loc = 169,column = '日付_168日目', value = None)
wb_5.insert(loc = 170,column = '日付_169日目', value = None)
wb_5.insert(loc = 171,column = '日付_170日目', value = None)
wb_5.insert(loc = 172,column = '日付_171日目', value = None)
wb_5.insert(loc = 173,column = '日付_172日目', value = None)
wb_5.insert(loc = 174,column = '日付_173日目', value = None)
wb_5.insert(loc = 175,column = '日付_174日目', value = None)
wb_5.insert(loc = 176,column = '日付_175日目', value = None)
wb_5.insert(loc = 177,column = '日付_176日目', value = None)
wb_5.insert(loc = 178,column = '日付_177日目', value = None)
wb_5.insert(loc = 179,column = '日付_178日目', value = None)
wb_5.insert(loc = 180,column = '日付_179日目', value = None)
wb_5.insert(loc = 181,column = '日付_180日目', value = None)
wb_5.insert(loc = 182,column = '日付_181日目', value = None)
wb_5.insert(loc = 183,column = '日付_182日目', value = None)
wb_5.insert(loc = 184,column = '日付_183日目', value = None)
wb_5.insert(loc = 185,column = '日付_184日目', value = None)
wb_5.insert(loc = 186,column = '日付_185日目', value = None)
wb_5.insert(loc = 187,column = '日付_186日目', value = None)
wb_5.insert(loc = 188,column = '日付_187日目', value = None)
wb_5.insert(loc = 189,column = '日付_188日目', value = None)
wb_5.insert(loc = 190,column = '日付_189日目', value = None)
wb_5.insert(loc = 191,column = '日付_190日目', value = None)
wb_5.insert(loc = 192,column = '日付_191日目', value = None)
wb_5.insert(loc = 193,column = '日付_192日目', value = None)
wb_5.insert(loc = 194,column = '日付_193日目', value = None)
wb_5.insert(loc = 195,column = '日付_194日目', value = None)
wb_5.insert(loc = 196,column = '日付_195日目', value = None)
wb_5.insert(loc = 197,column = '日付_196日目', value = None)
wb_5.insert(loc = 198,column = '日付_197日目', value = None)
wb_5.insert(loc = 199,column = '日付_198日目', value = None)
wb_5.insert(loc = 200,column = '日付_199日目', value = None)
wb_5.insert(loc = 201,column = '日付_200日目', value = None)
wb_5.insert(loc = 202,column = '日付_201日目', value = None)
wb_5.insert(loc = 203,column = '日付_202日目', value = None)
wb_5.insert(loc = 204,column = '日付_203日目', value = None)
wb_5.insert(loc = 205,column = '日付_204日目', value = None)
wb_5.insert(loc = 206,column = '日付_205日目', value = None)
wb_5.insert(loc = 207,column = '日付_206日目', value = None)
wb_5.insert(loc = 208,column = '日付_207日目', value = None)
wb_5.insert(loc = 209,column = '日付_208日目', value = None)
wb_5.insert(loc = 210,column = '日付_209日目', value = None)
wb_5.insert(loc = 211,column = '日付_210日目', value = None)
wb_5.insert(loc = 212,column = '日付_211日目', value = None)
wb_5.insert(loc = 213,column = '日付_212日目', value = None)
wb_5.insert(loc = 214,column = '日付_213日目', value = None)
wb_5.insert(loc = 215,column = '日付_214日目', value = None)
wb_5.insert(loc = 216,column = '日付_215日目', value = None)
wb_5.insert(loc = 217,column = '日付_216日目', value = None)
wb_5.insert(loc = 218,column = '日付_217日目', value = None)
wb_5.insert(loc = 219,column = '日付_218日目', value = None)
wb_5.insert(loc = 220,column = '日付_219日目', value = None)
wb_5.insert(loc = 221,column = '日付_220日目', value = None)
wb_5.insert(loc = 222,column = '日付_221日目', value = None)
wb_5.insert(loc = 223,column = '日付_222日目', value = None)
wb_5.insert(loc = 224,column = '日付_223日目', value = None)
wb_5.insert(loc = 225,column = '日付_224日目', value = None)
wb_5.insert(loc = 226,column = '日付_225日目', value = None)
wb_5.insert(loc = 227,column = '日付_226日目', value = None)
wb_5.insert(loc = 228,column = '日付_227日目', value = None)
wb_5.insert(loc = 229,column = '日付_228日目', value = None)
wb_5.insert(loc = 230,column = '日付_229日目', value = None)
wb_5.insert(loc = 231,column = '日付_230日目', value = None)
wb_5.insert(loc = 304,column = '内容_73日目', value = None)
wb_5.insert(loc = 305,column = '内容_74日目', value = None)
wb_5.insert(loc = 306,column = '内容_75日目', value = None)
wb_5.insert(loc = 307,column = '内容_76日目', value = None)
wb_5.insert(loc = 308,column = '内容_77日目', value = None)
wb_5.insert(loc = 309,column = '内容_78日目', value = None)
wb_5.insert(loc = 310,column = '内容_79日目', value = None)
wb_5.insert(loc = 311,column = '内容_80日目', value = None)
wb_5.insert(loc = 312,column = '内容_81日目', value = None)
wb_5.insert(loc = 313,column = '内容_82日目', value = None)
wb_5.insert(loc = 314,column = '内容_83日目', value = None)
wb_5.insert(loc = 315,column = '内容_84日目', value = None)
wb_5.insert(loc = 316,column = '内容_85日目', value = None)
wb_5.insert(loc = 317,column = '内容_86日目', value = None)
wb_5.insert(loc = 318,column = '内容_87日目', value = None)
wb_5.insert(loc = 319,column = '内容_88日目', value = None)
wb_5.insert(loc = 320,column = '内容_89日目', value = None)
wb_5.insert(loc = 321,column = '内容_90日目', value = None)
wb_5.insert(loc = 322,column = '内容_91日目', value = None)
wb_5.insert(loc = 323,column = '内容_92日目', value = None)
wb_5.insert(loc = 324,column = '内容_93日目', value = None)
wb_5.insert(loc = 325,column = '内容_94日目', value = None)
wb_5.insert(loc = 326,column = '内容_95日目', value = None)
wb_5.insert(loc = 327,column = '内容_96日目', value = None)
wb_5.insert(loc = 328,column = '内容_97日目', value = None)
wb_5.insert(loc = 329,column = '内容_98日目', value = None)
wb_5.insert(loc = 330,column = '内容_99日目', value = None)
wb_5.insert(loc = 331,column = '内容_100日目', value = None)
wb_5.insert(loc = 332,column = '内容_101日目', value = None)
wb_5.insert(loc = 333,column = '内容_102日目', value = None)
wb_5.insert(loc = 334,column = '内容_103日目', value = None)
wb_5.insert(loc = 335,column = '内容_104日目', value = None)
wb_5.insert(loc = 336,column = '内容_105日目', value = None)
wb_5.insert(loc = 337,column = '内容_106日目', value = None)
wb_5.insert(loc = 338,column = '内容_107日目', value = None)
wb_5.insert(loc = 339,column = '内容_108日目', value = None)
wb_5.insert(loc = 340,column = '内容_109日目', value = None)
wb_5.insert(loc = 341,column = '内容_110日目', value = None)
wb_5.insert(loc = 342,column = '内容_111日目', value = None)
wb_5.insert(loc = 343,column = '内容_112日目', value = None)
wb_5.insert(loc = 344,column = '内容_113日目', value = None)
wb_5.insert(loc = 345,column = '内容_114日目', value = None)
wb_5.insert(loc = 346,column = '内容_115日目', value = None)
wb_5.insert(loc = 347,column = '内容_116日目', value = None)
wb_5.insert(loc = 348,column = '内容_117日目', value = None)
wb_5.insert(loc = 349,column = '内容_118日目', value = None)
wb_5.insert(loc = 350,column = '内容_119日目', value = None)
wb_5.insert(loc = 351,column = '内容_120日目', value = None)
wb_5.insert(loc = 352,column = '内容_121日目', value = None)
wb_5.insert(loc = 353,column = '内容_122日目', value = None)
wb_5.insert(loc = 354,column = '内容_123日目', value = None)
wb_5.insert(loc = 355,column = '内容_124日目', value = None)
wb_5.insert(loc = 356,column = '内容_125日目', value = None)
wb_5.insert(loc = 357,column = '内容_126日目', value = None)
wb_5.insert(loc = 358,column = '内容_127日目', value = None)
wb_5.insert(loc = 359,column = '内容_128日目', value = None)
wb_5.insert(loc = 360,column = '内容_129日目', value = None)
wb_5.insert(loc = 361,column = '内容_130日目', value = None)
wb_5.insert(loc = 362,column = '内容_131日目', value = None)
wb_5.insert(loc = 363,column = '内容_132日目', value = None)
wb_5.insert(loc = 364,column = '内容_133日目', value = None)
wb_5.insert(loc = 365,column = '内容_134日目', value = None)
wb_5.insert(loc = 366,column = '内容_135日目', value = None)
wb_5.insert(loc = 367,column = '内容_136日目', value = None)
wb_5.insert(loc = 368,column = '内容_137日目', value = None)
wb_5.insert(loc = 369,column = '内容_138日目', value = None)
wb_5.insert(loc = 370,column = '内容_139日目', value = None)
wb_5.insert(loc = 371,column = '内容_140日目', value = None)
wb_5.insert(loc = 372,column = '内容_141日目', value = None)
wb_5.insert(loc = 373,column = '内容_142日目', value = None)
wb_5.insert(loc = 374,column = '内容_143日目', value = None)
wb_5.insert(loc = 375,column = '内容_144日目', value = None)
wb_5.insert(loc = 376,column = '内容_145日目', value = None)
wb_5.insert(loc = 377,column = '内容_146日目', value = None)
wb_5.insert(loc = 378,column = '内容_147日目', value = None)
wb_5.insert(loc = 379,column = '内容_148日目', value = None)
wb_5.insert(loc = 380,column = '内容_149日目', value = None)
wb_5.insert(loc = 381,column = '内容_150日目', value = None)
wb_5.insert(loc = 382,column = '内容_151日目', value = None)
wb_5.insert(loc = 383,column = '内容_152日目', value = None)
wb_5.insert(loc = 384,column = '内容_153日目', value = None)
wb_5.insert(loc = 385,column = '内容_154日目', value = None)
wb_5.insert(loc = 386,column = '内容_155日目', value = None)
wb_5.insert(loc = 387,column = '内容_156日目', value = None)
wb_5.insert(loc = 388,column = '内容_157日目', value = None)
wb_5.insert(loc = 389,column = '内容_158日目', value = None)
wb_5.insert(loc = 390,column = '内容_159日目', value = None)
wb_5.insert(loc = 391,column = '内容_160日目', value = None)
wb_5.insert(loc = 392,column = '内容_161日目', value = None)
wb_5.insert(loc = 393,column = '内容_162日目', value = None)
wb_5.insert(loc = 394,column = '内容_163日目', value = None)
wb_5.insert(loc = 395,column = '内容_164日目', value = None)
wb_5.insert(loc = 396,column = '内容_165日目', value = None)
wb_5.insert(loc = 397,column = '内容_166日目', value = None)
wb_5.insert(loc = 398,column = '内容_167日目', value = None)
wb_5.insert(loc = 399,column = '内容_168日目', value = None)
wb_5.insert(loc = 400,column = '内容_169日目', value = None)
wb_5.insert(loc = 401,column = '内容_170日目', value = None)
wb_5.insert(loc = 402,column = '内容_171日目', value = None)
wb_5.insert(loc = 403,column = '内容_172日目', value = None)
wb_5.insert(loc = 404,column = '内容_173日目', value = None)
wb_5.insert(loc = 405,column = '内容_174日目', value = None)
wb_5.insert(loc = 406,column = '内容_175日目', value = None)
wb_5.insert(loc = 407,column = '内容_176日目', value = None)
wb_5.insert(loc = 408,column = '内容_177日目', value = None)
wb_5.insert(loc = 409,column = '内容_178日目', value = None)
wb_5.insert(loc = 410,column = '内容_179日目', value = None)
wb_5.insert(loc = 411,column = '内容_180日目', value = None)
wb_5.insert(loc = 412,column = '内容_181日目', value = None)
wb_5.insert(loc = 413,column = '内容_182日目', value = None)
wb_5.insert(loc = 414,column = '内容_183日目', value = None)
wb_5.insert(loc = 415,column = '内容_184日目', value = None)
wb_5.insert(loc = 416,column = '内容_185日目', value = None)
wb_5.insert(loc = 417,column = '内容_186日目', value = None)
wb_5.insert(loc = 418,column = '内容_187日目', value = None)
wb_5.insert(loc = 419,column = '内容_188日目', value = None)
wb_5.insert(loc = 420,column = '内容_189日目', value = None)
wb_5.insert(loc = 421,column = '内容_190日目', value = None)
wb_5.insert(loc = 422,column = '内容_191日目', value = None)
wb_5.insert(loc = 423,column = '内容_192日目', value = None)
wb_5.insert(loc = 424,column = '内容_193日目', value = None)
wb_5.insert(loc = 425,column = '内容_194日目', value = None)
wb_5.insert(loc = 426,column = '内容_195日目', value = None)
wb_5.insert(loc = 427,column = '内容_196日目', value = None)
wb_5.insert(loc = 428,column = '内容_197日目', value = None)
wb_5.insert(loc = 429,column = '内容_198日目', value = None)
wb_5.insert(loc = 430,column = '内容_199日目', value = None)
wb_5.insert(loc = 431,column = '内容_200日目', value = None)
wb_5.insert(loc = 432,column = '内容_201日目', value = None)
wb_5.insert(loc = 433,column = '内容_202日目', value = None)
wb_5.insert(loc = 434,column = '内容_203日目', value = None)
wb_5.insert(loc = 435,column = '内容_204日目', value = None)
wb_5.insert(loc = 436,column = '内容_205日目', value = None)
wb_5.insert(loc = 437,column = '内容_206日目', value = None)
wb_5.insert(loc = 438,column = '内容_207日目', value = None)
wb_5.insert(loc = 439,column = '内容_208日目', value = None)
wb_5.insert(loc = 440,column = '内容_209日目', value = None)
wb_5.insert(loc = 441,column = '内容_210日目', value = None)
wb_5.insert(loc = 442,column = '内容_211日目', value = None)
wb_5.insert(loc = 443,column = '内容_212日目', value = None)
wb_5.insert(loc = 444,column = '内容_213日目', value = None)
wb_5.insert(loc = 445,column = '内容_214日目', value = None)
wb_5.insert(loc = 446,column = '内容_215日目', value = None)
wb_5.insert(loc = 447,column = '内容_216日目', value = None)
wb_5.insert(loc = 448,column = '内容_217日目', value = None)
wb_5.insert(loc = 449,column = '内容_218日目', value = None)
wb_5.insert(loc = 450,column = '内容_219日目', value = None)
wb_5.insert(loc = 451,column = '内容_220日目', value = None)
wb_5.insert(loc = 452,column = '内容_221日目', value = None)
wb_5.insert(loc = 453,column = '内容_222日目', value = None)
wb_5.insert(loc = 454,column = '内容_223日目', value = None)
wb_5.insert(loc = 455,column = '内容_224日目', value = None)
wb_5.insert(loc = 456,column = '内容_225日目', value = None)
wb_5.insert(loc = 457,column = '内容_226日目', value = None)
wb_5.insert(loc = 458,column = '内容_227日目', value = None)
wb_5.insert(loc = 459,column = '内容_228日目', value = None)
wb_5.insert(loc = 460,column = '内容_229日目', value = None)
wb_5.insert(loc = 461,column = '内容_230日目', value = None)
wb_5.insert(loc = 532,column = '担当_71日目', value = None)
wb_5.insert(loc = 533,column = '担当_72日目', value = None)
wb_5.insert(loc = 534,column = '担当_73日目', value = None)
wb_5.insert(loc = 535,column = '担当_74日目', value = None)
wb_5.insert(loc = 536,column = '担当_75日目', value = None)
wb_5.insert(loc = 537,column = '担当_76日目', value = None)
wb_5.insert(loc = 538,column = '担当_77日目', value = None)
wb_5.insert(loc = 539,column = '担当_78日目', value = None)
wb_5.insert(loc = 540,column = '担当_79日目', value = None)
wb_5.insert(loc = 541,column = '担当_80日目', value = None)
wb_5.insert(loc = 542,column = '担当_81日目', value = None)
wb_5.insert(loc = 543,column = '担当_82日目', value = None)
wb_5.insert(loc = 544,column = '担当_83日目', value = None)
wb_5.insert(loc = 545,column = '担当_84日目', value = None)
wb_5.insert(loc = 546,column = '担当_85日目', value = None)
wb_5.insert(loc = 547,column = '担当_86日目', value = None)
wb_5.insert(loc = 548,column = '担当_87日目', value = None)
wb_5.insert(loc = 549,column = '担当_88日目', value = None)
wb_5.insert(loc = 550,column = '担当_89日目', value = None)
wb_5.insert(loc = 551,column = '担当_90日目', value = None)
wb_5.insert(loc = 552,column = '担当_91日目', value = None)
wb_5.insert(loc = 553,column = '担当_92日目', value = None)
wb_5.insert(loc = 554,column = '担当_93日目', value = None)
wb_5.insert(loc = 555,column = '担当_94日目', value = None)
wb_5.insert(loc = 556,column = '担当_95日目', value = None)
wb_5.insert(loc = 557,column = '担当_96日目', value = None)
wb_5.insert(loc = 558,column = '担当_97日目', value = None)
wb_5.insert(loc = 559,column = '担当_98日目', value = None)
wb_5.insert(loc = 560,column = '担当_99日目', value = None)
wb_5.insert(loc = 561,column = '担当_100日目', value = None)
wb_5.insert(loc = 562,column = '担当_101日目', value = None)
wb_5.insert(loc = 563,column = '担当_102日目', value = None)
wb_5.insert(loc = 564,column = '担当_103日目', value = None)
wb_5.insert(loc = 565,column = '担当_104日目', value = None)
wb_5.insert(loc = 566,column = '担当_105日目', value = None)
wb_5.insert(loc = 567,column = '担当_106日目', value = None)
wb_5.insert(loc = 568,column = '担当_107日目', value = None)
wb_5.insert(loc = 569,column = '担当_108日目', value = None)
wb_5.insert(loc = 570,column = '担当_109日目', value = None)
wb_5.insert(loc = 571,column = '担当_110日目', value = None)
wb_5.insert(loc = 572,column = '担当_111日目', value = None)
wb_5.insert(loc = 573,column = '担当_112日目', value = None)
wb_5.insert(loc = 574,column = '担当_113日目', value = None)
wb_5.insert(loc = 575,column = '担当_114日目', value = None)
wb_5.insert(loc = 576,column = '担当_115日目', value = None)
wb_5.insert(loc = 577,column = '担当_116日目', value = None)
wb_5.insert(loc = 578,column = '担当_117日目', value = None)
wb_5.insert(loc = 579,column = '担当_118日目', value = None)
wb_5.insert(loc = 580,column = '担当_119日目', value = None)
wb_5.insert(loc = 581,column = '担当_120日目', value = None)
wb_5.insert(loc = 582,column = '担当_121日目', value = None)
wb_5.insert(loc = 583,column = '担当_122日目', value = None)
wb_5.insert(loc = 584,column = '担当_123日目', value = None)
wb_5.insert(loc = 585,column = '担当_124日目', value = None)
wb_5.insert(loc = 586,column = '担当_125日目', value = None)
wb_5.insert(loc = 587,column = '担当_126日目', value = None)
wb_5.insert(loc = 588,column = '担当_127日目', value = None)
wb_5.insert(loc = 589,column = '担当_128日目', value = None)
wb_5.insert(loc = 590,column = '担当_129日目', value = None)
wb_5.insert(loc = 591,column = '担当_130日目', value = None)
wb_5.insert(loc = 592,column = '担当_131日目', value = None)
wb_5.insert(loc = 593,column = '担当_132日目', value = None)
wb_5.insert(loc = 594,column = '担当_133日目', value = None)
wb_5.insert(loc = 595,column = '担当_134日目', value = None)
wb_5.insert(loc = 596,column = '担当_135日目', value = None)
wb_5.insert(loc = 597,column = '担当_136日目', value = None)
wb_5.insert(loc = 598,column = '担当_137日目', value = None)
wb_5.insert(loc = 599,column = '担当_138日目', value = None)
wb_5.insert(loc = 600,column = '担当_139日目', value = None)
wb_5.insert(loc = 601,column = '担当_140日目', value = None)
wb_5.insert(loc = 602,column = '担当_141日目', value = None)
wb_5.insert(loc = 603,column = '担当_142日目', value = None)
wb_5.insert(loc = 604,column = '担当_143日目', value = None)
wb_5.insert(loc = 605,column = '担当_144日目', value = None)
wb_5.insert(loc = 606,column = '担当_145日目', value = None)
wb_5.insert(loc = 607,column = '担当_146日目', value = None)
wb_5.insert(loc = 608,column = '担当_147日目', value = None)
wb_5.insert(loc = 609,column = '担当_148日目', value = None)
wb_5.insert(loc = 610,column = '担当_149日目', value = None)
wb_5.insert(loc = 611,column = '担当_150日目', value = None)
wb_5.insert(loc = 612,column = '担当_151日目', value = None)
wb_5.insert(loc = 613,column = '担当_152日目', value = None)
wb_5.insert(loc = 614,column = '担当_153日目', value = None)
wb_5.insert(loc = 615,column = '担当_154日目', value = None)
wb_5.insert(loc = 616,column = '担当_155日目', value = None)
wb_5.insert(loc = 617,column = '担当_156日目', value = None)
wb_5.insert(loc = 618,column = '担当_157日目', value = None)
wb_5.insert(loc = 619,column = '担当_158日目', value = None)
wb_5.insert(loc = 620,column = '担当_159日目', value = None)
wb_5.insert(loc = 621,column = '担当_160日目', value = None)
wb_5.insert(loc = 622,column = '担当_161日目', value = None)
wb_5.insert(loc = 623,column = '担当_162日目', value = None)
wb_5.insert(loc = 624,column = '担当_163日目', value = None)
wb_5.insert(loc = 625,column = '担当_164日目', value = None)
wb_5.insert(loc = 626,column = '担当_165日目', value = None)
wb_5.insert(loc = 627,column = '担当_166日目', value = None)
wb_5.insert(loc = 628,column = '担当_167日目', value = None)
wb_5.insert(loc = 629,column = '担当_168日目', value = None)
wb_5.insert(loc = 630,column = '担当_169日目', value = None)
wb_5.insert(loc = 631,column = '担当_170日目', value = None)
wb_5.insert(loc = 632,column = '担当_171日目', value = None)
wb_5.insert(loc = 633,column = '担当_172日目', value = None)
wb_5.insert(loc = 634,column = '担当_173日目', value = None)
wb_5.insert(loc = 635,column = '担当_174日目', value = None)
wb_5.insert(loc = 636,column = '担当_175日目', value = None)
wb_5.insert(loc = 637,column = '担当_176日目', value = None)
wb_5.insert(loc = 638,column = '担当_177日目', value = None)
wb_5.insert(loc = 639,column = '担当_178日目', value = None)
wb_5.insert(loc = 640,column = '担当_179日目', value = None)
wb_5.insert(loc = 641,column = '担当_180日目', value = None)
wb_5.insert(loc = 642,column = '担当_181日目', value = None)
wb_5.insert(loc = 643,column = '担当_182日目', value = None)
wb_5.insert(loc = 644,column = '担当_183日目', value = None)
wb_5.insert(loc = 645,column = '担当_184日目', value = None)
wb_5.insert(loc = 646,column = '担当_185日目', value = None)
wb_5.insert(loc = 647,column = '担当_186日目', value = None)
wb_5.insert(loc = 648,column = '担当_187日目', value = None)
wb_5.insert(loc = 649,column = '担当_188日目', value = None)
wb_5.insert(loc = 650,column = '担当_189日目', value = None)
wb_5.insert(loc = 651,column = '担当_190日目', value = None)
wb_5.insert(loc = 652,column = '担当_191日目', value = None)
wb_5.insert(loc = 653,column = '担当_192日目', value = None)
wb_5.insert(loc = 654,column = '担当_193日目', value = None)
wb_5.insert(loc = 655,column = '担当_194日目', value = None)
wb_5.insert(loc = 656,column = '担当_195日目', value = None)
wb_5.insert(loc = 657,column = '担当_196日目', value = None)
wb_5.insert(loc = 658,column = '担当_197日目', value = None)
wb_5.insert(loc = 659,column = '担当_198日目', value = None)
wb_5.insert(loc = 660,column = '担当_199日目', value = None)
wb_5.insert(loc = 661,column = '担当_200日目', value = None)
wb_5.insert(loc = 662,column = '担当_201日目', value = None)
wb_5.insert(loc = 663,column = '担当_202日目', value = None)
wb_5.insert(loc = 664,column = '担当_203日目', value = None)
wb_5.insert(loc = 665,column = '担当_204日目', value = None)
wb_5.insert(loc = 666,column = '担当_205日目', value = None)
wb_5.insert(loc = 667,column = '担当_206日目', value = None)
wb_5.insert(loc = 668,column = '担当_207日目', value = None)
wb_5.insert(loc = 669,column = '担当_208日目', value = None)
wb_5.insert(loc = 670,column = '担当_209日目', value = None)
wb_5.insert(loc = 671,column = '担当_210日目', value = None)
wb_5.insert(loc = 672,column = '担当_211日目', value = None)
wb_5.insert(loc = 673,column = '担当_212日目', value = None)

# データフレームを結合
wb = pd.concat([wb_1, wb_2, wb_3, wb_4, wb_5])

wb = wb.rename(columns={'NO.' : 'No'})

# 不要な値を含む行を削除
wb = wb.dropna(subset=['No'])
wb = wb.drop_duplicates(subset='No', keep='last')

# 関数db_insertの呼び出し
db_insert(wb, excel_db_map)