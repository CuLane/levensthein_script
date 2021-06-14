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
    "RITM0010464",
    "RITM0011926",
    "RITM0012780",
    "RITM0013041",
    "RITM0014681",
    "RITM0019670",
    "RITM0021039",
    "RITM0024527",
    "RITM0026023",
    "RITM0027832",
    "RITM0029601",
    "RITM0031999",
    "RITM0033520",
    "RITM0036979",
    "RITM0038974",
    "RITM0039532",
    "RITM0040837",
    "RITM0042486",
    "RITM0042687",
    "RITM0047309",
    "RITM0047478",
    "RITM0048470",
    "RITM0049029",
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
                {"tenant": tenant, "landlord": landlord, "match_type": "tenant email"}
            )
            continue
        elif lower_strip(tenant["Requested for"]) == lower_strip(
            landlord["Requested for"]
        ):
            matched_list.append(
                {
                    "tenant": tenant,
                    "landlord": landlord,
                    "match_type": "requested for",
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
                    "match_type": "address line 1",
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
                    "match_type": "landlord name",
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
                    "match_type": "landlord email",
                }
            )
            continue
        elif (
            landlord_domain not in COMMON_DOMAINS and tenant_domain not in COMMON_DOMAINS
            and landlord_domain != "Empty"
            and tenant_landlord_domain == landlord_domain
            and landlord_tenant_domain == tenant_domain
        ):
            matched_list.append(
                {"tenant": tenant, "landlord": landlord, "match_type": "t + ll email domain"}
            )
            continue


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
        lower_strip(match["tenant"]["Requested for"]),
        lower_strip(match["landlord"]["Requested for"]),
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
        "Tenant Name (Tenant)": lower_strip(match["tenant"]["Requested for"]),
        "Tenant Name (LL)": lower_strip(match["landlord"]["Requested for"]),
        "Tenant Name Comparison": requested_for,
        #
        "Tenant Address 1 + 2 (Tenant)": f"{lower_strip(match['tenant']['Address line 1'])} {lower_strip(match['tenant']['Address line 2'])}",
        "Tenant Address 1 + 2 (LL)": f"{lower_strip(match['landlord']['Address line 1'])} {lower_strip(match['landlord']['Address line 2'])}",
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
        "Landlord Name (Tenant)": lower_strip(match["tenant"]["Landlord Name"]),
        "Landlord Name (LL)": lower_strip(match["landlord"]["Landlord Name"]),
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


# make domain check check for both landlord domain and tenant domain
# sort by ritm tenant number, then by overall ratio
# Make the output file a command line argument
# Update columns
# verify zip fields
