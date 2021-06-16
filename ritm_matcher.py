#!/usr/bin/env python
# coding: utf-8

import os

import Levenshtein as lev
import pandas as pd


RATIO_THRESHHOLD = 0.7  # TODO Update threshold
CURRENT_DIRECTORY = os.getcwd()
OUTPUT_FILENAME = "matches.xlsx"
LANDLORDS_FILENAME = "landlords.xlsx"
TENANTS_FILENAME = "tenants.xlsx"
LANDLORDS_PATH = f"{CURRENT_DIRECTORY}/{LANDLORDS_FILENAME}"
TENANTS_PATH = f"{CURRENT_DIRECTORY}/{TENANTS_FILENAME}"
COMMON_DOMAINS = [
    "gmail.com",
    "yahoo.com",
    "aol.com",
    "hotmail.com",
    "outlook.com",
    "Empty",
]
# Paste unmatched RITMs here.
UNMATCHED_RITMS = [
"RITM0010108",
"RITM0010123",
"RITM0010131",
"RITM0010138",
"RITM0010146",
"RITM0010153",
"RITM0010173",
"RITM0010175",
"RITM0010190",
"RITM0010210",
"RITM0010229",
"RITM0010239",
"RITM0010301",
"RITM0010311",
"RITM0010368",
"RITM0010401",
"RITM0010494",
"RITM0010526",
"RITM0010609",
"RITM0010628",
"RITM0010721",
"RITM0010724",
"RITM0010751",
"RITM0010761",
"RITM0010765",
"RITM0010773",
"RITM0010849",
"RITM0010859",
"RITM0010877",
"RITM0010975",
"RITM0011015",
"RITM0011028",
"RITM0011120",
"RITM0011127",
"RITM0011161",
"RITM0011164",
"RITM0011173",
"RITM0011176",
"RITM0011197",
"RITM0011203",
"RITM0011204",
"RITM0011261",
"RITM0011287",
"RITM0011300",
"RITM0011306",
"RITM0011319",
"RITM0011365",
"RITM0011375",
"RITM0011418",
"RITM0011421",
"RITM0011449",
"RITM0011480",
"RITM0011493",
"RITM0011495",
"RITM0011504",
"RITM0011530",
"RITM0011666",
"RITM0011708",
"RITM0011733",
"RITM0011835",
"RITM0011895",
"RITM0011904",
"RITM0011944",
"RITM0011965",
"RITM0012023",
"RITM0012056",
"RITM0012131",
"RITM0012165",
"RITM0012198",
"RITM0012283",
"RITM0012374",
"RITM0012388",
"RITM0012453",
"RITM0012454",
"RITM0012459",
"RITM0012476",
"RITM0012481",
"RITM0012486",
"RITM0012525",
"RITM0012550",
"RITM0012573",
"RITM0012592",
"RITM0012633",
"RITM0012647",
"RITM0012652",
"RITM0012653",
"RITM0012654",
"RITM0012767",
"RITM0012830",
"RITM0012836",
"RITM0012871",
"RITM0012884",
"RITM0012933",
"RITM0013096",
"RITM0013102",
"RITM0013152",
"RITM0013173",
"RITM0013182",
"RITM0013220",
"RITM0013257",
"RITM0013315",
"RITM0013331",
"RITM001350118",
"RITM0013516",
"RITM0013534",
"RITM0013538",
"RITM0013542",
"RITM0013550",
"RITM0013551",
"RITM0013563",
"RITM0013573",
"RITM0013574",
"RITM0013594",
"RITM0013602",
"RITM0013663",
"RITM0013706",
"RITM0013793",
"RITM0013895",
"RITM0013932",
"RITM0013986",
"RITM0014023",
"RITM0014062",
"RITM0014069",
"RITM0014176",
"RITM0014187",
"RITM0014204",
"RITM0014263",
"RITM0014302",
"RITM0014328",
"RITM0014368",
"RITM0014410",
"RITM0014444",
"RITM0014455",
"RITM0014568",
"RITM0014663",
"RITM0014703",
"RITM0014773",
"RITM0014831",
"RITM0014848",
"RITM0014850",
"RITM0014874",
"RITM0014879",
"RITM0014891",
"RITM0014904",
"RITM0014943",
"RITM0014967",
"RITM0015019",
"RITM0015085",
"RITM0015202",
"RITM0015210",
"RITM0015258",
"RITM0015267",
"RITM0015276",
"RITM0015333",
"RITM0015343",
"RITM0015375",
"RITM0015442",
"RITM0015481",
"RITM0015566",
"RITM0015579",
"RITM0015684",
"RITM0015715",
"RITM0015744",
"RITM0015865",
"RITM0015874",
"RITM0015899",
"RITM0015967",
"RITM0016013",
"RITM0016016",
"RITM0016033",
"RITM0016128",
"RITM0016152",
"RITM0016258",
"RITM0016312",
"RITM0016367",
"RITM0016582",
"RITM0016592",
"RITM0016673",
"RITM0016747",
"RITM0016826",
"RITM0016896",
"RITM0016938",
"RITM0016997",
"RITM0017057",
"RITM0017246",
"RITM0017250",
"RITM0017496",
"RITM0017542",
"RITM0017548",
"RITM0017569",
"RITM0017603",
"RITM0017618",
"RITM0017629",
"RITM0017630",
"RITM0017700",
"RITM0017701",
"RITM0017744",
"RITM0017780",
"RITM0017826",
"RITM0017831",
"RITM0017998",
"RITM0018025",
"RITM0018040",
"RITM0018080",
"RITM0018081",
"RITM0018098",
"RITM0018111",
"RITM0018179",
"RITM0018293",
"RITM0018331",
"RITM0018395",
"RITM0018602",
"RITM0018612",
"RITM0018623",
"RITM0018733",
"RITM0018899",
"RITM0018904",
"RITM0018907",
"RITM0018948",
"RITM0019140",
"RITM0019276",
"RITM0019282",
"RITM0019371",
"RITM0019463",
"RITM0019482",
"RITM0019631",
"RITM0019702",
"RITM0019713",
"RITM0019733",
"RITM0019818",
"RITM0019819",
"RITM0019978",
"RITM0020046",
"RITM0020147",
"RITM0020310",
"RITM0020387",
"RITM0020388",
"RITM0020533",
"RITM0020541",
"RITM0020656",
"RITM0020732",
"RITM0020796",
"RITM0020807",
"RITM0020810",
"RITM0020838",
"RITM0021009",
"RITM0021019",
"RITM0021091",
"RITM0021270",
"RITM0021283",
"RITM0021306",
"RITM0021368",
"RITM0021448",
"RITM0021728",
"RITM0021780",
"RITM0021902",
"RITM0021997",
"RITM0022162",
"RITM0022235",
"RITM0022263",
"RITM0022432",
"RITM0022484",
"RITM0022538",
"RITM0022548",
"RITM0022662",
"RITM0022676",
"RITM0022687",
"RITM0022719",
"RITM0022729",
"RITM0022740",
"RITM0022782",
"RITM0022790",
"RITM0022892",
"RITM0022955",
"RITM0023036",
"RITM0023169",
"RITM0023174",
"RITM0023192",
"RITM0023295",
"RITM0023366",
"RITM0023440",
"RITM0023442",
"RITM0023619",
"RITM0023620",
"RITM0023688",
"RITM0023691",
"RITM0023938",
"RITM0023940",
"RITM0024035",
"RITM0024042",
"RITM0024081",
"RITM0024100",
"RITM0024128",
"RITM0024153",
"RITM0024160",
"RITM0024170",
"RITM0024191",
"RITM0024196",
"RITM0024231",
"RITM0024416",
"RITM0024559",
"RITM0024571",
"RITM0024607",
"RITM0024771",
"RITM0024828",
"RITM0024836",
"RITM0024919",
"RITM0025010",
"RITM0025093",
"RITM0025218",
"RITM0025223",
"RITM0025379",
"RITM0025449",
"RITM0025527",
"RITM0025545",
"RITM0025648",
"RITM0025729",
"RITM0025760",
"RITM0025801",
"RITM0025920",
"RITM0025989",
"RITM0026179",
"RITM0026391",
"RITM0026548",
"RITM0026569",
"RITM0026631",
"RITM0026813",
"RITM0026833",
"RITM0026853",
"RITM0026861",
"RITM0026931",
"RITM0027193",
"RITM0027214",
"RITM0027228",
"RITM0027253",
"RITM0027284",
"RITM0027323",
"RITM0027383",
"RITM0027729",
"RITM0027732",
"RITM0027820",
"RITM0027866",
"RITM0027926",
"RITM0028027",
"RITM0028073",
"RITM0028317",
"RITM0028356",
"RITM0028363",
"RITM0028381",
"RITM0028390",
"RITM0028585",
"RITM0028722",
"RITM0028840",
"RITM0028963",
"RITM0029088",
"RITM0029364",
"RITM0029372",
"RITM0029375",
"RITM0029437",
"RITM0029707",
"RITM0029781",
"RITM0029829",
"RITM0029951",
"RITM0029983",
"RITM0030051",
"RITM0030085",
"RITM0030115",
"RITM0030151",
"RITM0030223",
"RITM0030231",
"RITM0030335",
"RITM0030361",
"RITM0030425",
"RITM0030455",
"RITM0030508",
"RITM0030601",
"RITM0030731",
"RITM0030805",
"RITM0030811",
"RITM0030816",
"RITM0030921",
"RITM0030957",
"RITM0031070",
"RITM0031222",
"RITM0031298",
"RITM0031360",
"RITM0031425",
"RITM0031544",
"RITM0031578",
"RITM0031604",
"RITM0031615",
"RITM0031671",
"RITM0031678",
"RITM0031705",
"RITM0031795",
"RITM0032143",
"RITM0032174",
"RITM0032175",
"RITM0032202",
"RITM0032239",
"RITM0032309",
"RITM0032317",
"RITM0032353",
"RITM0032366",
"RITM0032450",
"RITM0032454",
"RITM0032514",
"RITM0032547",
"RITM0032579",
"RITM0032619",
"RITM0032672",
"RITM0032819",
"RITM0032823",
"RITM0032883",
"RITM0032886",
"RITM0032902",
"RITM0032914",
"RITM0032920",
"RITM0032999",
"RITM0033164",
"RITM0033173",
"RITM0033334",
"RITM0033352",
"RITM0033353",
"RITM0033387",
"RITM0033395",
"RITM0033419",
"RITM0033460",
"RITM0033491",
"RITM0033527",
"RITM0033545",
"RITM0033791",
"RITM0033820",
"RITM0033869",
"RITM0033916",
"RITM0034017",
"RITM0034234",
"RITM0034327",
"RITM0034337",
"RITM0034364",
"RITM0034427",
"RITM0034696",
"RITM0034834",
"RITM0034848",
"RITM0034901",
"RITM0034990",
"RITM0035135",
"RITM0035379",
"RITM0035392",
"RITM0035414",
"RITM0035426",
"RITM0035483",
"RITM0035504",
"RITM0035642",
"RITM0035755",
"RITM0035799",
"RITM0035877",
"RITM0036043",
"RITM0036101",
"RITM0036265",
"RITM0036306",
"RITM0036320",
"RITM0036446",
"RITM0036453",
"RITM0036475",
"RITM0036489",
"RITM0036505",
"RITM0036512",
"RITM0036581",
"RITM0036586",
"RITM0036622",
"RITM0036669",
"RITM0036910",
"RITM0036958",
"RITM0036961",
"RITM0037144",
"RITM0037247",
"RITM0037254",
"RITM0037282",
"RITM0037411",
"RITM0037493",
"RITM0037530",
"RITM0037569",
"RITM0037604",
"RITM0037612",
"RITM0037684",
"RITM0037696",
"RITM0037747",
"RITM0037758",
"RITM0037823",
"RITM0037890",
"RITM0037951",
"RITM0038212",
"RITM0038345",
"RITM0038417",
"RITM0038624",
"RITM0038677",
"RITM0038749",
"RITM0038778",
"RITM0038837",
"RITM0038985",
"RITM0039066",
"RITM0039091",
"RITM0039169",
"RITM0039459",
"RITM0039489",
"RITM0039540",
"RITM0039633",
"RITM0039637",
"RITM0039644",
"RITM0039680",
"RITM0039721",
"RITM0039762",
"RITM0039823",
"RITM0039858",
"RITM0039896",
"RITM0039907",
"RITM0039922",
"RITM0039989",
"RITM0040011",
"RITM0040037",
"RITM0040063",
"RITM0040077",
"RITM0040198",
"RITM0040257",
"RITM0040317",
"RITM0040380",
"RITM0040392",
"RITM0040406",
"RITM0040817",
"RITM0040873",
"RITM0040910",
"RITM0040934",
"RITM0040944",
"RITM0041289",
"RITM0041317",
"RITM0041409",
"RITM0041451",
"RITM0041457",
"RITM0041458",
"RITM0041673",
"RITM0041733",
"RITM0041866",
"RITM0041890",
"RITM0041961",
"RITM0042194",
"RITM0042295",
"RITM0042358",
"RITM0042415",
"RITM0042524",
"RITM0042786",
"RITM0042852",
"RITM0042963",
"RITM0043123",
"RITM0043201",
"RITM0043291",
"RITM0043323",
"RITM0043424",
"RITM0043426",
"RITM0043441",
"RITM0043684",
"RITM0043891",
"RITM0043940",
"RITM0044179",
"RITM0044209",
"RITM0044360",
"RITM0044476",
"RITM0044578",
"RITM0044619",
"RITM0044685",
"RITM0044823",
"RITM0045034",
"RITM0045093",
"RITM0045129",
"RITM0045180",
"RITM0045188",
"RITM0045230",
"RITM0045358",
"RITM0045523",
"RITM0045576",
"RITM0045702",
"RITM0045744",
"RITM0045900",
"RITM0045901",
"RITM0045938",
"RITM0046085",
"RITM0046096",
"RITM0046134",
"RITM0046258",
"RITM0046580",
"RITM0046823",
"RITM0046830",
"RITM0046836",
"RITM0046884",
"RITM0046896",
"RITM0046969",
"RITM0046983",
"RITM0047067",
"RITM0047186",
"RITM0047207",
"RITM0047218",
"RITM0047242",
"RITM0047273",
"RITM0047324",
"RITM0047474",
"RITM0047559",
"RITM0047726",
"RITM0047857",
"RITM0047883",
"RITM0047944",
"RITM0047999",
"RITM0048002",
"RITM0048125",
"RITM0048178",
"RITM0048369",
"RITM0048557",
"RITM0048761",
"RITM0048876",
"RITM0048881",
"RITM0049077",
"RITM0049105",
"RITM0049241",
"RITM0049254",
"RITM0049501",
"RITM0049756",
"RITM0049786",
"RITM0049861",
"RITM0050145",
"RITM0050205",
]

# Load data
landlords = pd.read_excel(LANDLORDS_PATH)
landlords.fillna("Empty")
print(f"{len(landlords)} Unmatched Landlord RITMs")

tenants = pd.read_excel(TENANTS_PATH)
tenants.fillna("Empty")
print(f"{len(tenants)} Unmatched Tenant RITMs")

filtered_tenants = tenants[tenants.Number.isin(UNMATCHED_RITMS)]
print(f"Checking {len(filtered_tenants)} tenants for matches.")
matched_list = []
results = []


# Helper methods
def lower_strip(value):
    """
    Takes a value and returns it as a string,
    lowercased and stripped of extra whitespace
    """
    return str(value).lower().strip()


def color_red(val):
    """
    Takes a scalar and returns a string with
    the css property `'color: red'` for ratios
    greater than or equal to 0.7, black otherwise.
    """
    try:
        color = "red" if float(val) >= RATIO_THRESHHOLD and float(val) <= 1 else "black"
    except ValueError:
        color = "black"
    return "color: %s" % color


def if_float(match, ll_email_ratio):
    """
    returns the lev ratio unless landlord email isn't a string,
    then it returns 0
    """
    if not isinstance(match["tenant"]["Landlord Email"], str):
        return 0
    else:
        return ll_email_ratio


for i, landlord in landlords.iterrows():
    """
    Prematch by comparing lower cased and
    trimmed values.
    """
    landlord_domain = ""
    landlord_tenant_domain = ""  #
    if (i + 1) % 100 == 0:
        print(f"Checking landlord #{i + 1}")
    if str("@") in str(landlord["Landlord Email"]):
        landlord_domain = lower_strip(landlord["Landlord Email"]).split("@")[1]
    if str("@") in str(landlord["Tenant Email"]):
        landlord_tenant_domain = lower_strip(landlord["Tenant Email"]).split("@")[1]
        #
        #
    for ii, tenant in filtered_tenants.iterrows():
        tenant_landlord_domain = ""
        tenant_domain = ""  #
        if str("@") in str(tenant["Landlord Email"]):
            tenant_landlord_domain = lower_strip(tenant["Landlord Email"]).split("@")[1]
        if str("@") in str(tenant["Tenant Email"]):
            tenant_domain = lower_strip(tenant["Tenant Email"]).split("@")[1]

        if lower_strip(tenant["Tenant Email"]) == lower_strip(landlord["Tenant Email"]):
            matched_list.append(
                {"tenant": tenant, "landlord": landlord, "match_type": "Tenant Email"}
            )
            continue
        elif lower_strip(tenant["Requested for"]) ==  f"{lower_strip(landlord['Tenant first name'])} {lower_strip(landlord['Tenant last name'])}":
            matched_list.append(
                {
                    "tenant": tenant,
                    "landlord": landlord,
                    "match_type": "Tenant Name",
                }
            )
            continue
        elif (
            landlord_domain not in COMMON_DOMAINS
            and tenant_domain not in COMMON_DOMAINS
            and landlord_domain != "Empty"
            and tenant_landlord_domain == landlord_domain
            and landlord_tenant_domain == tenant_domain
        ):
            matched_list.append(
                {
                    "tenant": tenant,
                    "landlord": landlord,
                    "match_type": "t + ll email domain",
                }
            )
            continue
        elif lower_strip(tenant["Address line 1"]) == lower_strip(
            landlord["Address line 1"]
        ):
            matched_list.append(
                {
                    "tenant": tenant,
                    "landlord": landlord,
                    "match_type": "Address Line 1",
                }
            )
            continue
        elif lower_strip(tenant["Landlord Email"]) == lower_strip(
            landlord["Landlord Email"]
        ):
            matched_list.append(
                {
                    "tenant": tenant,
                    "landlord": landlord,
                    "match_type": "Landlord Email",
                }
            )
            continue
        elif lower_strip(tenant["Landlord Name"]) == lower_strip(
            landlord["Landlord Name"]
        ):
            matched_list.append(
                {
                    "tenant": tenant,
                    "landlord": landlord,
                    "match_type": "Landlord Name",
                }
            )
            continue


def remove_nan(value):
    if not isinstance(value, str) or value == 'nan':
        return ""
    else:
        return value


for match in matched_list:
    """
    Get Levenshtein Ratio for each match
    """
    tenant_email = lev.ratio(
        lower_strip(match["tenant"]["Tenant Email"]),
        lower_strip(match["landlord"]["Tenant Email"]),
    )
    address_line = lev.ratio(
        f"{lower_strip(match['tenant']['Address line 1'])} {lower_strip(match['tenant']['Address line 2'])}",
        f"{lower_strip(match['landlord']['Address line 1'])} {lower_strip(match['landlord']['Address line 2'])}",
    )
    requested_for = lev.ratio(
        remove_nan(lower_strip(match["tenant"]["Requested for"])),
        f"{remove_nan(lower_strip(match['landlord']['Tenant first name']))} {remove_nan(lower_strip(match['landlord']['Tenant last name']))}",
    )
    landlord_name = lev.ratio(
        lower_strip(match["tenant"]["Landlord Name"]),
        lower_strip(match["landlord"]["Landlord Name"]),
    )
    landlord_email = lev.ratio(
        lower_strip(match["tenant"]["Landlord Email"]),
        lower_strip(match["landlord"]["Landlord Email"]),
    )
    zip_code = lev.ratio(
        lower_strip(int(match["tenant"]["Zip Code"])),
        lower_strip(int(match["landlord"]["Zip Code"])),
    )
    ratios = {
        "Tenant RITM": match["tenant"]["Number"],
        "LL RITM": match["landlord"]["Number"],
        "Match Type": match["match_type"],
        #
        "Tenant Name (Tenant)": remove_nan(lower_strip(match["tenant"]["Requested for"])),
        "Tenant Name (LL)": f"{remove_nan(lower_strip(match['landlord']['Tenant first name']))} {remove_nan(lower_strip(match['landlord']['Tenant last name']))}",
        "Tenant Name Comparison": requested_for,
        #
        "Tenant Address 1 + 2 (Tenant)": f"{lower_strip(match['tenant']['Address line 1'])} {remove_nan(lower_strip(match['tenant']['Address line 2']))}",
        "Tenant Address 1 + 2 (LL)": f"{lower_strip(match['landlord']['Address line 1'])} {remove_nan(lower_strip(match['landlord']['Address line 2']))}",
        "Address Line Comparison": address_line,
        #
        "Tenant Zip Code (Tenant)": lower_strip(int(match["tenant"]["Zip Code"])),
        "Tenant Zip Code (LL)": lower_strip(int(match["landlord"]["Zip Code"])),
        "Tenant Zip Code Comparison": zip_code,
        #
        "Tenant Email (Tenant)": match["tenant"]["Tenant Email"],
        "Tenant Email (LL)": match["landlord"]["Tenant Email"],
        "Tenant Email Comparison": tenant_email,
        #
        "Landlord Name (Tenant)": remove_nan(lower_strip(match["tenant"]["Landlord Name"])),
        "Landlord Name (LL)": remove_nan(lower_strip(match["landlord"]["Landlord Name"])),
        "Landlord Name Comparison": landlord_name,
        #
        "Landlord Email (Tenant)": match["tenant"]["Landlord Email"],
        "Landlord Email (LL)": match["landlord"]["Landlord Email"],
        "Landlord Email Comparison": if_float(match, landlord_email),
        #
        "Comparison Average": "{:.2f}".format(
            (
                landlord_email
                + tenant_email
                + landlord_name
                + requested_for
                + address_line
                + zip_code
            )
            / 6
        ),
    }
    results.append(ratios)

print(f"{len(results)} potential matches found")

results_df = pd.DataFrame(results)
results_df = results_df.style.applymap(color_red)
results_df.to_excel(OUTPUT_FILENAME)

print(f"{OUTPUT_FILENAME} file created at {CURRENT_DIRECTORY}.")
print("All done! ♪┏(°.°)┛┗(°.°)┓┗(°.°)┛┏(°.°)┓ ♪")

# TODO Make the output filename a command line argument.

