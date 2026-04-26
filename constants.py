from openpyxl.styles import Border, Side

FILE_PATH = "final Apprisal.xlsm"
LOGO_PATH = "logo.png"
SETTINGS_FILE = "app_settings.json"
TRIAL_FILE = "trial_users.json"
USERS_FILE = "users.json"

MONTH_MAP = {
    "يناير":"January","فبراير":"February","مارس":"March",
    "أبريل":"April","مايو":"May","يونيو":"June",
    "يوليو":"July","أغسطس":"August","سبتمبر":"September",
    "أكتوبر":"October","نوفمبر":"November","ديسمبر":"December",
    "Jan":"January","Feb":"February","Mar":"March","Apr":"April",
    "May":"May","Jun":"June","Jul":"July","Aug":"August",
    "Sep":"September","Oct":"October","Nov":"November","Dec":"December",
}
MONTHS_AR = ["يناير","فبراير","مارس","أبريل","مايو","يونيو",
             "يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر"]
MONTHS_EN = ["January","February","March","April","May","June",
             "July","August","September","October","November","December"]
MONTHS_SHORT = ["Jan","Feb","Mar","Apr","May","Jun",
                "Jul","Aug","Sep","Oct","Nov","Dec"]

PERSONAL_KPIS = [
    "الالتزام بساعات الدوام اليومي ومكان العمل",
    "الاهتمام بالمظهر العام والهندام والمحافظة على نظافة وترتيب",
    "مبادر وخلاق ومتعاون ويحافظ على علاقات ايجابية",
    "يتحمل ضغط العمل ولا يتذمر عند طلب اعمال اضافية",
    "يتحلى بالامانة والمصداقية ولا يفشي اسرار العمل او الزملاء",
]
PERSONAL_WEIGHT = 4

DARK="1F3864"; MID="2E75B6"; LBLUE="BDD7EE"; ORANGE="ED7D31"
YELLOW="FFD966"; LGRAY="F2F2F2"; GREEN_BG="E2EFDA"
RED_BG="FCE4D6"; WHITE="FFFFFF"; CREAM="FFFBF0"

thick_s = Side(style="medium", color="1F3864")
thin_s = Side(style="thin", color="AAAAAA")
OUTER_B = Border(left=thick_s, right=thick_s, top=thick_s, bottom=thick_s)
INNER_B = Border(left=thin_s, right=thin_s, top=thin_s, bottom=thin_s)
ROW_B = Border(left=thick_s, right=thick_s, top=thin_s, bottom=thin_s)

