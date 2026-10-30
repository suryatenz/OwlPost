"""
add_new_depts.py
Adds faculty from new departments not yet covered:
  - UH College of Optometry (Vision Science + Clinical Sciences)
  - UH College of Pharmacy (PHOP dept - health outcomes & policy)
  - UH Hilton College of Global Hospitality Leadership
  - UH CLASS - Psychology
  - UH CLASS - Political Science
  - UH CLASS - Sociology
  - UH Graduate College of Social Work
  - UH-Downtown - Computer Science & Engineering Technology
  - UH-Downtown - Mathematics & Statistics
  - UH-Clear Lake - Computer Science
  - UH-Clear Lake - Data Science
  - UH-Clear Lake - Computer Information Systems
  - UH-Clear Lake - Mathematics
  - UH-Clear Lake - Statistics
  + Deans/Directors for all new colleges

Skips any email already present in professors.xlsx.
"""

import openpyxl
from openpyxl.styles import Alignment, Border, Side
import os

EXCEL = os.path.join(os.path.dirname(os.path.abspath(__file__)), "professors.xlsx")

# ── Data ──────────────────────────────────────────────────────────────────────

FACULTY = [

    # ── UH College of Optometry - Vision Science Faculty ──────────────────────
    {"name": "Muayyad Al-Ubaidi",       "email": "malubaid@central.uh.edu",    "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "Raymond Applegate",        "email": "rapplegate@uh.edu",          "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "Julia Benoit",             "email": "jsbenoit@central.uh.edu",    "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "Daniel Coates",            "email": "drcoates@central.uh.edu",    "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "Vivien Coulson-Thomas",    "email": "vjcoulso@central.uh.edu",    "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "Vallabh Das",              "email": "vdas@central.uh.edu",        "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "Luca Della Santina",       "email": "ldellasantina@uh.edu",       "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "Tarsis Ferreira",          "email": "tgferrei@central.uh.edu",    "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "Laura Frishman",           "email": "lfrishman@uh.edu",           "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "David Hildebrand",         "email": "dhildebrand@uh.edu",         "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "Nange Jin",                "email": "njin@uh.edu",                "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "Kirill Larin",             "email": "klarin@uh.edu",              "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "Jason Marsack",            "email": "jmarsack@central.uh.edu",    "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "Anna Matynia",             "email": "amatynia@central.uh.edu",    "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "Kimberly Meier",           "email": "kmeier@central.uh.edu",      "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "Muna Naash",               "email": "mnaash@central.uh.edu",      "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "John O'Brien",             "email": "jobrien3@uh.edu",            "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "Nimesh Patel",             "email": "npatel@central.uh.edu",      "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "Jason Porter",             "email": "jporter@central.uh.edu",     "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "Guoting Qin",              "email": "gqin@central.uh.edu",        "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "Rachel Redfern",           "email": "rredfer2@central.uh.edu",    "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "Christophe Ribelayga",     "email": "cpribela@central.uh.edu",    "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "Alexander Schill",         "email": "awschill@central.uh.edu",    "dept": "Optometry - Vision Science",  "rank": "Faculty"},
    {"name": "Geunyoung Yoon",           "email": "gyoon2@uh.edu",              "dept": "Optometry - Vision Science",  "rank": "Faculty"},

    # ── UH College of Optometry - Clinical Sciences Faculty ───────────────────
    {"name": "Andrew Archila",           "email": "amarchila@uh.edu",           "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Jan Bergmanson",           "email": "jbergmanson@uh.edu",         "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "David Berntsen",           "email": "dberntsen@uh.edu",           "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Moriah Chandler",          "email": "machand2@central.uh.edu",    "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Han Cheng",                "email": "hcheng@central.uh.edu",      "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Jennifer Deakins",         "email": "jdeakins@central.uh.edu",    "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "David Dinh",               "email": "dtdinh@central.uh.edu",      "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Meron Fitta",              "email": "mafitta@central.uh.edu",     "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Janet Garza",              "email": "jgarza6@central.uh.edu",     "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Amber Gaume",              "email": "agaume@uh.edu",              "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Marcus Gonzales",          "email": "mbgonzal@central.uh.edu",    "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Anita Gostovic",           "email": "aticak@central.uh.edu",      "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Beth Harper",              "email": "bharper@central.uh.edu",     "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Wendy Harrison",           "email": "wwharris@central.uh.edu",    "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Ralph Herring",            "email": "rjherring@uh.edu",           "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Hong Ho",                  "email": "hdho@central.uh.edu",        "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Kassaundra Johnston",      "email": "kjohnsto@central.uh.edu",    "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Michael Kalloniatis",      "email": "mkalloni@central.uh.edu",    "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Andrew Kemp",              "email": "apkemp@central.uh.edu",      "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Tonya Ketcham",            "email": "tketcham@central.uh.edu",    "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Jung-Sun Kim",             "email": "jkim88@uh.edu",              "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Carol Lakkis",             "email": "clakkis@central.uh.edu",     "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Kimberly Lambreghts",      "email": "klambreghts@uh.edu",         "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Anna-Kaye Logan",          "email": "amlogan@central.uh.edu",     "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Danica Marrelli",          "email": "dmarrelli@uh.edu",           "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Taylor Marsh",             "email": "tkmarsh@central.uh.edu",     "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Swati Modi",               "email": "scmodi@central.uh.edu",      "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Courtney Morrow",          "email": "cfmorrow@central.uh.edu",    "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Chi Nguyen",               "email": "lcnguye2@central.uh.edu",    "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Amber Nichols",            "email": "ahicks4@uh.edu",             "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Lloyd Pate Jr",            "email": "ldpate@central.uh.edu",      "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Maureen Plaumann",         "email": "mdplauma@central.uh.edu",    "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Kathryn Richdale",         "email": "richdale@central.uh.edu",    "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Eric Ritchey",             "email": "erritche@central.uh.edu",    "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Kaitlyn Sapoznik",         "email": "kasapoznik@uh.edu",          "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Diane Sayah",              "email": "dnsayah@central.uh.edu",     "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Krishna Surapaneni",       "email": "krsurapa@central.uh.edu",    "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Jennifer Tasca",           "email": "jmtasca@central.uh.edu",     "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Anuradha Veerappan",       "email": "aveerapp@central.uh.edu",    "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Maria Walker",             "email": "mkwalker@uh.edu",            "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Ashley Wallace-Tucker",    "email": "awallace@central.uh.edu",    "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Mary Kate Walters",        "email": "mksain@central.uh.edu",      "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Samantha Wang",            "email": "swang50@uh.edu",             "dept": "Optometry - Clinical Sciences","rank": "Faculty"},
    {"name": "Joe Wheat",                "email": "jlwheat@central.uh.edu",     "dept": "Optometry - Clinical Sciences","rank": "Faculty"},

    # ── UH College of Pharmacy - PHOP (Health Outcomes & Policy) ─────────────
    {"name": "Susan Abughosh",           "email": "smabughosh@uh.edu",          "dept": "Pharmacy - Health Outcomes & Policy", "rank": "Faculty"},
    {"name": "Rajender Aparasu",         "email": "rraparasu@uh.edu",           "dept": "Pharmacy - Health Outcomes & Policy", "rank": "Chair"},
    {"name": "Hua Chen",                 "email": "hchen20@uh.edu",             "dept": "Pharmacy - Health Outcomes & Policy", "rank": "Faculty"},
    {"name": "Ekere James Essien",       "email": "ejessien@uh.edu",            "dept": "Pharmacy - Health Outcomes & Policy", "rank": "Faculty"},
    {"name": "Jieni Li",                 "email": "jli87@central.uh.edu",       "dept": "Pharmacy - Health Outcomes & Policy", "rank": "Faculty"},
    {"name": "Xi Lu",                    "email": "xlu20@uh.edu",               "dept": "Pharmacy - Health Outcomes & Policy", "rank": "Faculty"},
    {"name": "Dipali Rinker",            "email": "dvrinker@central.uh.edu",    "dept": "Pharmacy - Health Outcomes & Policy", "rank": "Faculty"},
    {"name": "Sujit Sansgiry",           "email": "SSansgiry@uh.edu",           "dept": "Pharmacy - Health Outcomes & Policy", "rank": "Faculty"},
    {"name": "J. Douglas Thornton",      "email": "jdthornt@central.uh.edu",    "dept": "Pharmacy - Health Outcomes & Policy", "rank": "Faculty"},
    {"name": "Moosa Tatar",              "email": "mtatar@central.uh.edu",      "dept": "Pharmacy - Health Outcomes & Policy", "rank": "Faculty"},
    {"name": "Tiansheng Wang",           "email": "tianwang@uh.edu",            "dept": "Pharmacy - Health Outcomes & Policy", "rank": "Faculty"},

    # ── UH Hilton College of Global Hospitality Leadership ────────────────────
    {"name": "Dennis Reynolds",          "email": "der@uh.edu",                 "dept": "Hilton College",              "rank": "Dean"},
    {"name": "Ki-Joon Back",             "email": "kback@uh.edu",               "dept": "Hilton College",              "rank": "Faculty"},
    {"name": "Mary Dawson",              "email": "mdawson@uh.edu",             "dept": "Hilton College",              "rank": "Faculty"},
    {"name": "Cristian Morosan",         "email": "cmorosan@uh.edu",            "dept": "Hilton College",              "rank": "Faculty"},
    {"name": "Agnes L. DeFranco",        "email": "adefranco@uh.edu",           "dept": "Hilton College",              "rank": "Faculty"},
    {"name": "Priyanko Guchait",         "email": "pguchait@uh.edu",            "dept": "Hilton College",              "rank": "Faculty"},
    {"name": "Juan M. Madera",           "email": "jmmadera@uh.edu",            "dept": "Hilton College",              "rank": "Faculty"},
    {"name": "Akash Grover",             "email": "akgrover@uh.edu",            "dept": "Hilton College",              "rank": "Faculty"},
    {"name": "Yoon Koh",                 "email": "ykoh@uh.edu",                "dept": "Hilton College",              "rank": "Faculty"},
    {"name": "Minwoo Lee",               "email": "mlee37@uh.edu",              "dept": "Hilton College",              "rank": "Faculty"},
    {"name": "Jaewook Kim",              "email": "jkim65@uh.edu",              "dept": "Hilton College",              "rank": "Faculty"},
    {"name": "Ningqiao Li",              "email": "nli12@cougarnet.uh.edu",     "dept": "Hilton College",              "rank": "Faculty"},
    {"name": "Jay Neal",                 "email": "jneal2@uh.edu",              "dept": "Hilton College",              "rank": "Faculty"},
    {"name": "D. Christopher Taylor",    "email": "dctaylor@uh.edu",            "dept": "Hilton College",              "rank": "Faculty"},
    {"name": "Tucker Johnson",           "email": "tajohnson7@uh.edu",          "dept": "Hilton College",              "rank": "Faculty"},
    {"name": "Sujata Sirsat",            "email": "sasirsat@uh.edu",            "dept": "Hilton College",              "rank": "Faculty"},

    # ── UH CLASS - Psychology ─────────────────────────────────────────────────
    {"name": "Candice A. Alfano",        "email": "caalfano@uh.edu",            "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Julia Babcock",            "email": "jbabcock@uh.edu",            "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Johanna Bick",             "email": "jrbick@uh.edu",              "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Jeremy Borjon",            "email": "jiborjon@uh.edu",            "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Bradley Brummel",          "email": "bjbrummel@uh.edu",           "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "James Campion",            "email": "jcampion@uh.edu",            "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Coleen D. Carlson",        "email": "ccarlson@uh.edu",            "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Paul T. Cirino",           "email": "pcirino@uh.edu",             "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Kelly T. Cosgrove",        "email": "kcosgrove@uh.edu",           "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Rodica Damian",            "email": "ridamian@uh.edu",            "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Jaye L. Derrick",          "email": "jlderrick@uh.edu",           "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Abdo Elnakouri",           "email": "aelnakou@central.uh.edu",    "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Adam K. Fetterman",        "email": "akfetterman@uh.edu",         "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Jack M. Fletcher",         "email": "jackfletcher@uh.edu",        "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Donald J. Foss",           "email": "dfoss@uh.edu",               "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "David J. Francis",         "email": "dfrancis@uh.edu",            "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Matthew W. Gallagher",     "email": "mwgallagher@uh.edu",         "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Jason W. Griffin",         "email": "jasongriffin@central.uh.edu","dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Arturo E. Hernandez",      "email": "aehernandez@uh.edu",         "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "C. Raymond Knee",          "email": "knee@uh.edu",                "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Therese A. Kosten",        "email": "takosten@uh.edu",            "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Keke Lai",                 "email": "klai@central.uh.edu",        "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "J. Leigh Leasure",         "email": "jlleasure@uh.edu",           "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Kaylee Litson",            "email": "klitson@central.uh.edu",     "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Luis D. Medina",           "email": "LDMedina2@uh.edu",           "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Paras Mehta",              "email": "pmehta4@uh.edu",             "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Clayton Neighbors",        "email": "cneighbors@uh.edu",          "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Chelsea Ratcliff",         "email": "cgratcliff@central.uh.edu",  "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Denise Reyes",             "email": "dlreyes2@uh.edu",            "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Isaac Sabat",              "email": "isabat@uh.edu",              "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Carla Sharp",              "email": "csharp2@uh.edu",             "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Lynne Steinberg",          "email": "lsteinberg@uh.edu",          "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Amanda C. Venta",          "email": "aventa@uh.edu",              "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "John P. Vincent",          "email": "jvincent@uh.edu",            "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Michael W. Williams",      "email": "MWwilliams2@uh.edu",         "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "L. A. Witt",               "email": "witt@uh.edu",                "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Steven Paul Woods",        "email": "spwoods@uh.edu",             "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Hanako Yoshida",           "email": "yoshida@uh.edu",             "dept": "CLASS - Psychology",          "rank": "Faculty"},
    {"name": "Michael J. Zvolensky",     "email": "mjzvolen@central.uh.edu",    "dept": "CLASS - Psychology",          "rank": "Faculty"},

    # ── UH CLASS - Political Science ──────────────────────────────────────────
    {"name": "Max Allamong",             "email": "mballamong@uh.edu",          "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Eduardo Aleman",           "email": "ealeman2@uh.edu",            "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Allison Archer",           "email": "amarcher3@uh.edu",           "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Alex Badas",               "email": "abadas@uh.edu",              "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Tanya Bagashka",           "email": "tbagashka@uh.edu",           "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Scott J. Basinger",        "email": "sjbasinger@uh.edu",          "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Royce Carroll",            "email": "racarroll@uh.edu",           "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Jason P. Casellas",        "email": "jcasellas@uh.edu",           "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Tyson Chatagnier",         "email": "jtchatagnier@uh.edu",        "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Jennifer Clark",           "email": "jclar2@central.uh.edu",      "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Hayley M. Cohen",          "email": "hmcohen@uh.edu",             "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Cyrus Contractor",         "email": "ccontractor@uh.edu",         "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Jeronimo Cortina",         "email": "jcortina@central.uh.edu",    "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Joshua Freed",             "email": "jafreed2@uh.edu",            "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Tobias Heinrich",          "email": "theinrich@uh.edu",           "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Libby Jenke",              "email": "ljenke@uh.edu",              "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Michael Kistner",          "email": "mkistner@central.uh.edu",    "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Charles Nathan",           "email": "cnathan@uh.edu",             "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Brandon Rottinghaus",      "email": "bjrottinghaus@uh.edu",       "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Susan Scarrow",            "email": "sscarrow@uh.edu",            "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Boris Shor",               "email": "bshor@central.uh.edu",       "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Michael J. Soules",        "email": "msoules@uh.edu",             "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Lydia B. Tiede",           "email": "lbtiede@uh.edu",             "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Ling Zhu",                 "email": "lzhu4@central.uh.edu",       "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Zachary Zwald",            "email": "zjzwald@central.uh.edu",     "dept": "CLASS - Political Science",   "rank": "Faculty"},
    {"name": "Jessica Gottlieb",         "email": "jagottlieb@uh.edu",          "dept": "CLASS - Political Science",   "rank": "Faculty"},

    # ── UH CLASS - Sociology ──────────────────────────────────────────────────
    {"name": "Kathryn Freeman Anderson", "email": "kateanderson@uh.edu",        "dept": "CLASS - Sociology",           "rank": "Faculty"},
    {"name": "Amanda Baumle",            "email": "akbaumle@uh.edu",            "dept": "CLASS - Sociology",           "rank": "Faculty"},
    {"name": "Stella Grigorian",         "email": "sgrigorian@uh.edu",          "dept": "CLASS - Sociology",           "rank": "Faculty"},
    {"name": "Tracy Xavia Karner",       "email": "txkarner@uh.edu",            "dept": "CLASS - Sociology",           "rank": "Faculty"},
    {"name": "Sheila M. Katz",           "email": "smkatz@uh.edu",              "dept": "CLASS - Sociology",           "rank": "Faculty"},
    {"name": "Samantha Kwan",            "email": "sskwan@uh.edu",              "dept": "CLASS - Sociology",           "rank": "Faculty"},
    {"name": "Neema Langa",              "email": "nmlanga@uh.edu",             "dept": "CLASS - Sociology",           "rank": "Faculty"},
    {"name": "Shayne Lee",               "email": "slee3@uh.edu",               "dept": "CLASS - Sociology",           "rank": "Faculty"},
    {"name": "Maria A. Monserud",        "email": "mamonser@central.uh.edu",    "dept": "CLASS - Sociology",           "rank": "Faculty"},
    {"name": "Steven M. Nelson",         "email": "snelson4@uh.edu",            "dept": "CLASS - Sociology",           "rank": "Faculty"},
    {"name": "Carlee Purdum",            "email": "jcpurdum@uh.edu",            "dept": "CLASS - Sociology",           "rank": "Faculty"},
    {"name": "Pamela A. Quiroz",         "email": "paquiroz@uh.edu",            "dept": "CLASS - Sociology",           "rank": "Faculty"},
    {"name": "Scott V. Savage",          "email": "svsavage@uh.edu",            "dept": "CLASS - Sociology",           "rank": "Faculty"},
    {"name": "Zelma Oyarvide Tuthill",   "email": "zloyarvi@central.uh.edu",    "dept": "CLASS - Sociology",           "rank": "Faculty"},

    # ── UH Graduate College of Social Work ────────────────────────────────────
    {"name": "Samira Ali",               "email": "sbali3@central.uh.edu",      "dept": "Graduate College of Social Work","rank": "Faculty"},
    {"name": "Donna Amtsberg",           "email": "dkamtsbe@central.uh.edu",    "dept": "Graduate College of Social Work","rank": "Faculty"},
    {"name": "Jessica Bagneris",         "email": "jbagneri@central.uh.edu",    "dept": "Graduate College of Social Work","rank": "Faculty"},
    {"name": "Juan Barthelemy",          "email": "jbarthel@central.uh.edu",    "dept": "Graduate College of Social Work","rank": "Faculty"},
    {"name": "Jodi Berger Cardoso",      "email": "jabergercardoso@uh.edu",     "dept": "Graduate College of Social Work","rank": "Faculty"},
    {"name": "Sharon Borja",             "email": "sborja@central.uh.edu",      "dept": "Graduate College of Social Work","rank": "Faculty"},
    {"name": "Reiko Boyd",               "email": "rkboyd@central.uh.edu",      "dept": "Graduate College of Social Work","rank": "Faculty"},
    {"name": "Monit Cheung",             "email": "mcheung@uh.edu",             "dept": "Graduate College of Social Work","rank": "Faculty"},
    {"name": "Charday Foster",           "email": "cdfoster2@uh.edu",           "dept": "Graduate College of Social Work","rank": "Faculty"},
    {"name": "Robin Gearing",            "email": "rgearing@uh.edu",            "dept": "Graduate College of Social Work","rank": "Faculty"},
    {"name": "Shelley Gonzales",         "email": "sagonz12@central.uh.edu",    "dept": "Graduate College of Social Work","rank": "Faculty"},
    {"name": "Patrick Leung",            "email": "pleung@uh.edu",              "dept": "Graduate College of Social Work","rank": "Faculty"},
    {"name": "Virginia Lucas",           "email": "vlucas@central.uh.edu",      "dept": "Graduate College of Social Work","rank": "Faculty"},
    {"name": "Yvonne Mendoza",           "email": "yamendoz@central.uh.edu",    "dept": "Graduate College of Social Work","rank": "Faculty"},
    {"name": "Christina Miyawaki",       "email": "cemiyawa@central.uh.edu",    "dept": "Graduate College of Social Work","rank": "Faculty"},
    {"name": "Jamie Parker",             "email": "jparker@central.uh.edu",     "dept": "Graduate College of Social Work","rank": "Faculty"},
    {"name": "Steven L. Parks",          "email": "slparks@central.uh.edu",     "dept": "Graduate College of Social Work","rank": "Faculty"},
    {"name": "McClain Sampson",          "email": "mmsampson@uh.edu",           "dept": "Graduate College of Social Work","rank": "Faculty"},
    {"name": "Edward Scott",             "email": "edscottj@central.uh.edu",    "dept": "Graduate College of Social Work","rank": "Faculty"},
    {"name": "Isabel Torres",            "email": "istorres@uh.edu",            "dept": "Graduate College of Social Work","rank": "Faculty"},
    {"name": "Quenette Walton",          "email": "qwalton2@central.uh.edu",    "dept": "Graduate College of Social Work","rank": "Faculty"},
    {"name": "Tamika White",             "email": "tawhite5@central.uh.edu",    "dept": "Graduate College of Social Work","rank": "Faculty"},

    # ── UH-Downtown - Computer Science & Engineering Technology ───────────────
    {"name": "Vassilios Tzouanas",       "email": "tzouanasv@uhd.edu",          "dept": "UHD - Computer Science & Engineering Tech","rank": "Chair"},
    {"name": "Dvijesh Shastri",          "email": "shastrid@uhd.edu",           "dept": "UHD - Computer Science & Engineering Tech","rank": "Faculty"},
    {"name": "Henry Foust",              "email": "fousth@uhd.edu",             "dept": "UHD - Computer Science & Engineering Tech","rank": "Faculty"},
    {"name": "Cyril Harris",             "email": "harriscy@uhd.edu",           "dept": "UHD - Computer Science & Engineering Tech","rank": "Faculty"},
    {"name": "Joseph Kamto",             "email": "kamtoj@uhd.edu",             "dept": "UHD - Computer Science & Engineering Tech","rank": "Faculty"},
    {"name": "Hong Lin",                 "email": "linh@uhd.edu",               "dept": "UHD - Computer Science & Engineering Tech","rank": "Faculty"},
    {"name": "Azadeh Izadi",             "email": "izadia@uhd.edu",             "dept": "UHD - Computer Science & Engineering Tech","rank": "Faculty"},
    {"name": "Subash Pakhrin",           "email": "pakhrins@uhd.edu",           "dept": "UHD - Computer Science & Engineering Tech","rank": "Faculty"},
    {"name": "Kulwant Singh",            "email": "singhk@uhd.edu",             "dept": "UHD - Computer Science & Engineering Tech","rank": "Faculty"},
    {"name": "Benjamin Soibam",          "email": "soibamb@uhd.edu",            "dept": "UHD - Computer Science & Engineering Tech","rank": "Faculty"},
    {"name": "Ling Xu",                  "email": "xul@uhd.edu",                "dept": "UHD - Computer Science & Engineering Tech","rank": "Faculty"},
    {"name": "Emre Yilmaz",              "email": "yilmaze@uhd.edu",            "dept": "UHD - Computer Science & Engineering Tech","rank": "Faculty"},
    {"name": "Shengli Yuan",             "email": "yuans@uhd.edu",              "dept": "UHD - Computer Science & Engineering Tech","rank": "Faculty"},
    {"name": "Ting Zhang",               "email": "zhangtin@uhd.edu",           "dept": "UHD - Computer Science & Engineering Tech","rank": "Faculty"},

    # ── UH-Downtown - Mathematics & Statistics ────────────────────────────────
    {"name": "Susan L. Beane",           "email": "beanes@uhd.edu",             "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Bonnie Blumberg",          "email": "nadlerb@uhd.edu",            "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Trinh Bui",                "email": "buit@uhd.edu",               "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Aktug Cilekci",            "email": "cilekcia@uhd.edu",           "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Youn-Sha Chan",            "email": "chany@uhd.edu",              "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Ermelinda DeLaVina",       "email": "delavinae@uhd.edu",          "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Sangeeta Gad",             "email": "gads@uhd.edu",               "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Katarina Jegdic",          "email": "jegdick@uhd.edu",            "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Benny John",               "email": "johnb@uhd.edu",              "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Sergiy Koshkin",           "email": "koshkins@uhd.edu",           "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Mukesh Kumar",             "email": "kumarm@uhd.edu",             "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Jean Bernard Nganou",      "email": "nganouj@uhd.edu",            "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Mitsue Nakamura",          "email": "nakamuram@uhd.edu",          "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Vien Nguyen",              "email": "nguyenv@uhd.edu",            "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Ryan Pepper",              "email": "pepperr@uhd.edu",            "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Timothy Redl",             "email": "redlt@uhd.edu",              "dept": "UHD - Mathematics & Statistics","rank": "Chair"},
    {"name": "Plamen Simeonov",          "email": "simeonovp@uhd.edu",          "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Edwin Tecarro",            "email": "tecarroe@uhd.edu",           "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Shishen Xie",              "email": "xies@uhd.edu",               "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Jeong-Mi Yoon",            "email": "yoonj@uhd.edu",              "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Vasilis G. Zafiris",       "email": "zafirisv@uhd.edu",           "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Dexter Cahoy",             "email": "cahoyd@uhd.edu",             "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Wenlong Gong",             "email": "gongw@uhd.edu",              "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "J. Patrick King",          "email": "kingj@uhd.edu",              "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Kendra Mhoon",             "email": "mhoonk@uhd.edu",             "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Ha Nguyen",                "email": "nguyenha@uhd.edu",           "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Anna Rosenthal Simmons",   "email": "rosenthala@uhd.edu",         "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},
    {"name": "Katherine Shoemaker",      "email": "shoemakerk@uhd.edu",         "dept": "UHD - Mathematics & Statistics","rank": "Faculty"},

    # ── UH-Clear Lake - Computer Science ──────────────────────────────────────
    {"name": "Hisham Al-Mubaid",         "email": "hisham@uhcl.edu",            "dept": "UHCL - Computer Science",     "rank": "Chair"},
    {"name": "Krishani Abeysekera",      "email": "abeysekera@uhcl.edu",        "dept": "UHCL - Computer Science",     "rank": "Faculty"},
    {"name": "Ahmed Abukmail",           "email": "abukmail@uhcl.edu",          "dept": "UHCL - Computer Science",     "rank": "Faculty"},
    {"name": "Sayed Banawan",            "email": "banawan@uhcl.edu",           "dept": "UHCL - Computer Science",     "rank": "Faculty"},
    {"name": "Rishita Bansal",           "email": "bansal@uhcl.edu",            "dept": "UHCL - Computer Science",     "rank": "Faculty"},
    {"name": "Said Bettayeb",            "email": "bettayeb@uhcl.edu",          "dept": "UHCL - Computer Science",     "rank": "Faculty"},
    {"name": "Pradeep Buddharaju",       "email": "buddharaju@uhcl.edu",        "dept": "UHCL - Computer Science",     "rank": "Faculty"},
    {"name": "Khondker Shajadul Hasan",  "email": "hasank@uhcl.edu",            "dept": "UHCL - Computer Science",     "rank": "Faculty"},
    {"name": "Enamul Karim",             "email": "karim@uhcl.edu",             "dept": "UHCL - Computer Science",     "rank": "Faculty"},
    {"name": "Alfredo Perez-Davila",     "email": "perezd@uhcl.edu",            "dept": "UHCL - Computer Science",     "rank": "Faculty"},
    {"name": "Pavan Poudel",             "email": "poudel@uhcl.edu",            "dept": "UHCL - Computer Science",     "rank": "Faculty"},
    {"name": "Toneluh Yang",             "email": "yang@uhcl.edu",              "dept": "UHCL - Computer Science",     "rank": "Faculty"},
    {"name": "Kwok-Bun Yue",             "email": "yue@uhcl.edu",               "dept": "UHCL - Computer Science",     "rank": "Faculty"},

    # ── UH-Clear Lake - Data Science ──────────────────────────────────────────
    {"name": "Ross Niswanger",           "email": "niswanger@uhcl.edu",         "dept": "UHCL - Data Science",         "rank": "Faculty"},
    {"name": "Yalong Wu",                "email": "wuy@uhcl.edu",               "dept": "UHCL - Data Science",         "rank": "Faculty"},
    {"name": "Lijing Zhu",               "email": "zhul@uhcl.edu",              "dept": "UHCL - Data Science",         "rank": "Faculty"},

    # ── UH-Clear Lake - Computer Information Systems ──────────────────────────
    {"name": "Xinying Liu",              "email": "liux@uhcl.edu",              "dept": "UHCL - Computer Information Systems","rank": "Director"},

    # ── UH-Clear Lake - Mathematics ───────────────────────────────────────────
    {"name": "Nelson Carter",            "email": "carter@uhcl.edu",            "dept": "UHCL - Mathematics",          "rank": "Faculty"},
    {"name": "Manki Cho",                "email": "cho@uhcl.edu",               "dept": "UHCL - Mathematics",          "rank": "Faculty"},
    {"name": "Tom Fox",                  "email": "fox@uhcl.edu",               "dept": "UHCL - Mathematics",          "rank": "Chair"},
    {"name": "Michael Heerey",           "email": "heerey@uhcl.edu",            "dept": "UHCL - Mathematics",          "rank": "Faculty"},
    {"name": "Jingjing Ma",              "email": "ma@uhcl.edu",                "dept": "UHCL - Mathematics",          "rank": "Faculty"},
    {"name": "Zokhrab Mustafaev",        "email": "mustafaev@uhcl.edu",         "dept": "UHCL - Mathematics",          "rank": "Faculty"},
    {"name": "Teresa Rivas",             "email": "rivas@uhcl.edu",             "dept": "UHCL - Mathematics",          "rank": "Faculty"},
    {"name": "LieJune Shiau",            "email": "shiau@uhcl.edu",             "dept": "UHCL - Mathematics",          "rank": "Faculty"},
    {"name": "Yipeng Yang",              "email": "yangy@uhcl.edu",             "dept": "UHCL - Mathematics",          "rank": "Director"},

    # ── UH-Clear Lake - Statistics ────────────────────────────────────────────
    {"name": "Kalanka Jayalath",         "email": "jayalath@uhcl.edu",          "dept": "UHCL - Statistics",           "rank": "Director"},
    {"name": "Yingfu Li",                "email": "li@uhcl.edu",                "dept": "UHCL - Statistics",           "rank": "Faculty"},
    {"name": "Sadhana Weerasinghe",      "email": "weerasinghe@uhcl.edu",       "dept": "UHCL - Statistics",           "rank": "Faculty"},

    # ── Deans / Directors for new colleges ────────────────────────────────────
    {"name": "Hisham Al-Mubaid",         "email": "hisham@uhcl.edu",            "dept": "UHCL - Computing Sciences",   "rank": "Chair"},
    {"name": "Thomas Fox",               "email": "fox@uhcl.edu",               "dept": "UHCL - Mathematical Sciences","rank": "Chair"},
    {"name": "Kalanka Jayalath",         "email": "jayalath@uhcl.edu",          "dept": "UHCL - Statistics",           "rank": "Director"},
    {"name": "James Akif Uzman",         "email": "uzmana@uhd.edu",             "dept": "UHD - College of Sciences & Technology","rank": "Dean"},
    {"name": "Ermelinda DeLaVina",       "email": "delavinae@uhd.edu",          "dept": "UHD - Mathematics & Statistics","rank": "Associate Dean"},
]


# ── Spreadsheet helpers ───────────────────────────────────────────────────────

def main():
    wb = openpyxl.load_workbook(EXCEL)
    ws = wb.active

    headers = [str(c.value).strip() if c.value else "" for c in ws[1]]
    col = {h: i for i, h in enumerate(headers)}

    existing = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[col["Email"]]:
            existing.add(str(row[col["Email"]]).strip().lower())

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    added = 0
    skipped = 0
    seen_this_run = set()

    for f in FACULTY:
        email = f["email"].strip().lower()

        if email in existing or email in seen_this_run:
            print(f"  [=] Already present: {f['name']} <{email}>")
            skipped += 1
            continue

        values = {
            "Professor Name": f["name"],
            "Email":          email,
            "Department":     f["dept"],
            "AcadGroup":      "",
            "Rank":           f["rank"],
            "Sent":           "",
            "Stop":           "",
        }
        if "Replied" in col:
            values["Replied"] = ""

        r = ws.max_row + 1
        for i, key in enumerate(headers, 1):
            cell = ws.cell(row=r, column=i, value=values.get(key, ""))
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border = border

        existing.add(email)
        seen_this_run.add(email)
        added += 1
        print(f"  [+] {f['rank']:12s}  {f['name']} <{email}>")

    wb.save(EXCEL)

    print(f"\n{'='*60}")
    print(f"  New entries added    : {added}")
    print(f"  Already present      : {skipped}")
    print(f"  Saved to             : {EXCEL}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
