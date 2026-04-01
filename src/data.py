"""All static tournament data for FIFA World Cup 2026."""

from __future__ import annotations

GROUPS: dict[str, list[str]] = {
    "A": ["Mexico", "South Korea", "South Africa", "Czech Republic"],
    "B": ["Canada", "Switzerland", "Qatar", "Bosnia & Herzegovina"],
    "C": ["Brazil", "Morocco", "Scotland", "Haiti"],
    "D": ["USA", "Australia", "Paraguay", "Turkey"],
    "E": ["Germany", "Ecuador", "Ivory Coast", "Curaçao"],
    "F": ["Netherlands", "Japan", "Tunisia", "Sweden"],
    "G": ["Belgium", "Iran", "Egypt", "New Zealand"],
    "H": ["Spain", "Uruguay", "Cape Verde", "Saudi Arabia"],
    "I": ["France", "Senegal", "Norway", "Iraq"],
    "J": ["Argentina", "Algeria", "Austria", "Jordan"],
    "K": ["Portugal", "Colombia", "Uzbekistan", "DR Congo"],
    "L": ["England", "Croatia", "Panama", "Ghana"],
}

# (matchday, date, time_et, home, away, venue)
GROUP_MATCHES: dict[str, list[tuple]] = {
    "A": [
        (1, "Jun 11", "3 pm ET",  "Mexico",        "South Africa",        "Mexico City"),
        (1, "Jun 11", "10 pm ET", "South Korea",   "Czech Republic",      "Guadalajara"),
        (2, "Jun 18", "12 pm ET", "Czech Republic","South Africa",        "Atlanta"),
        (2, "Jun 18", "9 pm ET",  "Mexico",        "South Korea",         "Guadalajara"),
        (3, "Jun 24", "9 pm ET",  "Czech Republic","Mexico",              "Mexico City"),
        (3, "Jun 24", "9 pm ET",  "South Africa",  "South Korea",         "Monterrey"),
    ],
    "B": [
        (1, "Jun 12", "3 pm ET",  "Canada",              "Bosnia & Herzegovina", "Toronto"),
        (1, "Jun 13", "3 pm ET",  "Qatar",               "Switzerland",          "San Francisco"),
        (2, "Jun 18", "3 pm ET",  "Switzerland",         "Bosnia & Herzegovina", "Los Angeles"),
        (2, "Jun 18", "6 pm ET",  "Canada",              "Qatar",                "Vancouver"),
        (3, "Jun 24", "3 pm ET",  "Canada",              "Switzerland",          "Vancouver"),
        (3, "Jun 24", "3 pm ET",  "Bosnia & Herzegovina","Qatar",                "Seattle"),
    ],
    "C": [
        (1, "Jun 13", "6 pm ET",  "Brazil",   "Morocco",  "New York/NJ"),
        (1, "Jun 13", "9 pm ET",  "Haiti",    "Scotland", "Boston"),
        (2, "Jun 19", "3 pm ET",  "Scotland", "Morocco",  "Boston"),
        (2, "Jun 19", "9 pm ET",  "Brazil",   "Haiti",    "Philadelphia"),
        (3, "Jun 24", "6 pm ET",  "Scotland", "Brazil",   "Miami"),
        (3, "Jun 24", "6 pm ET",  "Morocco",  "Haiti",    "Atlanta"),
    ],
    "D": [
        (1, "Jun 12", "9 pm ET",  "USA",      "Paraguay", "Los Angeles"),
        (1, "Jun 13", "12 am ET", "Australia","Turkey",   "Vancouver"),
        (2, "Jun 19", "3 pm ET",  "USA",      "Australia","Seattle"),
        (2, "Jun 19", "12 am ET", "Turkey",   "Paraguay", "San Francisco"),
        (3, "Jun 25", "10 pm ET", "Turkey",   "USA",      "Los Angeles"),
        (3, "Jun 25", "10 pm ET", "Paraguay", "Australia","San Francisco"),
    ],
    "E": [
        (1, "Jun 14", "1 pm ET",  "Germany",     "Curaçao",     "Houston"),
        (1, "Jun 14", "7 pm ET",  "Ivory Coast", "Ecuador",     "Philadelphia"),
        (2, "Jun 20", "4 pm ET",  "Germany",     "Ivory Coast", "Toronto"),
        (2, "Jun 20", "8 pm ET",  "Ecuador",     "Curaçao",     "Kansas City"),
        (3, "Jun 25", "4 pm ET",  "Ecuador",     "Germany",     "New York/NJ"),
        (3, "Jun 25", "4 pm ET",  "Curaçao",     "Ivory Coast", "Philadelphia"),
    ],
    "F": [
        (1, "Jun 14", "4 pm ET",  "Netherlands","Japan",        "Dallas"),
        (1, "Jun 14", "10 pm ET", "Sweden",     "Tunisia",      "Monterrey"),
        (2, "Jun 20", "1 pm ET",  "Netherlands","Sweden",       "Houston"),
        (2, "Jun 20", "12 am ET", "Tunisia",    "Japan",        "Monterrey"),
        (3, "Jun 25", "7 pm ET",  "Japan",      "Sweden",       "Dallas"),
        (3, "Jun 25", "7 pm ET",  "Tunisia",    "Netherlands",  "Kansas City"),
    ],
    "G": [
        (1, "Jun 15", "3 pm ET",  "Belgium",     "Egypt",       "Seattle"),
        (1, "Jun 15", "9 pm ET",  "Iran",        "New Zealand", "Los Angeles"),
        (2, "Jun 21", "3 pm ET",  "Belgium",     "Iran",        "Los Angeles"),
        (2, "Jun 21", "9 pm ET",  "New Zealand", "Egypt",       "Vancouver"),
        (3, "Jun 26", "11 pm ET", "Egypt",       "Iran",        "Seattle"),
        (3, "Jun 26", "11 pm ET", "New Zealand", "Belgium",     "Vancouver"),
    ],
    "H": [
        (1, "Jun 15", "12 pm ET", "Spain",        "Cape Verde",   "Atlanta"),
        (1, "Jun 15", "6 pm ET",  "Saudi Arabia", "Uruguay",      "Miami"),
        (2, "Jun 21", "12 pm ET", "Spain",        "Saudi Arabia", "Atlanta"),
        (2, "Jun 21", "6 pm ET",  "Uruguay",      "Cape Verde",   "Miami"),
        (3, "Jun 26", "8 pm ET",  "Cape Verde",   "Saudi Arabia", "Houston"),
        (3, "Jun 26", "8 pm ET",  "Uruguay",      "Spain",        "Guadalajara"),
    ],
    "I": [
        (1, "Jun 16", "3 pm ET",  "France",  "Senegal", "New York/NJ"),
        (1, "Jun 16", "6 pm ET",  "Iraq",    "Norway",  "Boston"),
        (2, "Jun 22", "5 pm ET",  "France",  "Iraq",    "Philadelphia"),
        (2, "Jun 22", "8 pm ET",  "Norway",  "Senegal", "New York/NJ"),
        (3, "Jun 26", "3 pm ET",  "Norway",  "France",  "Boston"),
        (3, "Jun 26", "3 pm ET",  "Senegal", "Iraq",    "Toronto"),
    ],
    "J": [
        (1, "Jun 16", "9 pm ET",  "Argentina","Algeria",   "Kansas City"),
        (1, "Jun 16", "12 am ET", "Austria",  "Jordan",    "San Francisco"),
        (2, "Jun 22", "1 pm ET",  "Argentina","Austria",   "Dallas"),
        (2, "Jun 22", "11 pm ET", "Jordan",   "Algeria",   "San Francisco"),
        (3, "Jun 27", "10 pm ET", "Algeria",  "Austria",   "Kansas City"),
        (3, "Jun 27", "10 pm ET", "Jordan",   "Argentina", "Dallas"),
    ],
    "K": [
        (1, "Jun 17", "1 pm ET",    "Portugal",  "DR Congo",  "Houston"),
        (1, "Jun 17", "10 pm ET",   "Uzbekistan","Colombia",  "Mexico City"),
        (2, "Jun 23", "1 pm ET",    "Portugal",  "Uzbekistan","Houston"),
        (2, "Jun 23", "10 pm ET",   "DR Congo",  "Colombia",  "Guadalajara"),
        (3, "Jun 27", "7:30 pm ET", "Colombia",  "Portugal",  "Miami"),
        (3, "Jun 27", "7:30 pm ET", "DR Congo",  "Uzbekistan","Atlanta"),
    ],
    "L": [
        (1, "Jun 17", "4 pm ET", "England", "Croatia", "Dallas"),
        (1, "Jun 17", "7 pm ET", "Ghana",   "Panama",  "Toronto"),
        (2, "Jun 23", "4 pm ET", "England", "Ghana",   "Boston"),
        (2, "Jun 23", "7 pm ET", "Panama",  "Croatia", "Toronto"),
        (3, "Jun 27", "5 pm ET", "Panama",  "England", "New York/NJ"),
        (3, "Jun 27", "5 pm ET", "Croatia", "Ghana",   "Philadelphia"),
    ],
}

# Round of 32 matchups.
# team_l and team_r are either "GROUP_RANK" strings (resolved at runtime
# to formulas) or "3rd*" (manual input slots for best-3rd-place qualifiers).
R32_LEFT: list[tuple] = [
    ("A2", "B2",   "Jun 28", "Los Angeles"),
    ("C1", "F2",   "Jun 29", "Houston"),
    ("E1", "3rd*", "Jun 29", "Boston"),
    ("F1", "C2",   "Jun 29", "Monterrey"),
    ("E2", "I2",   "Jun 30", "Dallas"),
    ("I1", "3rd*", "Jun 30", "New York/NJ"),
    ("A1", "3rd*", "Jun 30", "Mexico City"),
    ("L1", "3rd*", "Jul 1",  "Atlanta"),
]
R32_RIGHT: list[tuple] = [
    ("G1",  "3rd*", "Jul 1",  "Seattle"),
    ("D1",  "3rd*", "Jul 1",  "San Francisco"),
    ("H1",  "J2",   "Jul 2",  "Los Angeles"),
    ("K2",  "L2",   "Jul 2",  "Toronto"),
    ("B1",  "3rd*", "Jul 2",  "Vancouver"),
    ("D2",  "G2",   "Jul 3",  "Dallas"),
    ("J1",  "H2",   "Jul 3",  "Miami"),
    ("K1",  "3rd*", "Jul 3",  "Kansas City"),
]

R16_LEFT: list[tuple] = [
    (0, 1,  "Jul 4", "Houston"),       # winners of R32-L[0] and R32-L[1]
    (2, 3,  "Jul 4", "Philadelphia"),
    (4, 5,  "Jul 5", "New York/NJ"),
    (6, 7,  "Jul 5", "Mexico City"),
]
R16_RIGHT: list[tuple] = [
    (0, 1,  "Jul 6", "Dallas"),        # winners of R32-R[0] and R32-R[1]
    (2, 3,  "Jul 6", "Seattle"),
    (4, 5,  "Jul 7", "Atlanta"),
    (6, 7,  "Jul 7", "Vancouver"),
]

QF_LEFT: list[tuple] = [
    (0, 1,  "Jul 9",  "Boston"),       # winners of R16-L[0] and R16-L[1]
    (2, 3,  "Jul 10", "Los Angeles"),
]
QF_RIGHT: list[tuple] = [
    (0, 1,  "Jul 11", "Miami"),
    (2, 3,  "Jul 11", "Kansas City"),
]

SF_LEFT  = (0, 1, "Jul 14", "Dallas")
SF_RIGHT = (0, 1, "Jul 15", "Atlanta")

FINAL    = ("Jul 19", "New York/NJ · MetLife Stadium")
THIRD    = ("Jul 18", "Miami")

REFERENCES = [
    (
        "FIFA Official",
        "https://www.fifa.com/en/tournaments/mens/worldcup/canadamexicousa2026"
        "/articles/match-schedule-fixtures-results-teams-stadiums",
        "Apr 1, 2026",
        "Official match schedule and fixtures",
    ),
    (
        "FOX Sports",
        "https://www.foxsports.com/stories/soccer/"
        "2026-world-cup-schedule-all-games-dates-matchups-how-watch",
        "Apr 1, 2026",
        "Full broadcast schedule with dates, times and venues",
    ),
    (
        "CNN Español",
        "https://cnnespanol.cnn.com/2026/04/01/deportes/"
        "grupos-zonas-copa-mundial-2026-orix",
        "Apr 1, 2026",
        "Complete groups and match calendar (post-playoff)",
    ),
    (
        "Marca",
        "https://www.marca.com/futbol/mundial/2026/04/01/"
        "estan-12-grupos-completos-mundial-falta-irak-bolivia.html",
        "Apr 1, 2026",
        "12 complete groups confirmed; Iraq & DR Congo qualify via inter-confederation playoff",
    ),
    (
        "Wikipedia – 2026 FIFA World Cup",
        "https://en.wikipedia.org/wiki/2026_FIFA_World_Cup",
        "Apr 1, 2026",
        "Tournament overview, format, host cities, qualification",
    ),
]
