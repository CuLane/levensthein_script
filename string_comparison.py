#!/usr/bin/env python
# coding: utf-8

import os
import pprint

import Levenshtein as lev
import pandas as pd

pp = pprint.PrettyPrinter(indent=4)

CURRENT_DIRECTORY = os.getcwd()
RATIO_THRESHHOLD = 0.7
OUTPUT_FILENAME = "matches.xlsx"
LANDLORDS_FILENAME = "landlords.xlsx"
TENANTS_FILENAME = "tenants.xlsx"
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
landlords_path = f"{CURRENT_DIRECTORY}/{LANDLORDS_FILENAME}"
landlords = pd.read_excel(landlords_path)
print(f"{len(landlords)} Unmatched Landlord RITMs")

tenants_path = f"{CURRENT_DIRECTORY}/{TENANTS_FILENAME}"
tenants = pd.read_excel(tenants_path)
print(f"{len(tenants)} Unmatched Tenant RITMs")

filtered_tenants = tenants[tenants.Number.isin(UNMATCHED_RITMS)]
matched_list = []
results = []


# Helper methods
def lower_strip(value):
    """
    Takes a value and returns it as a string,
    lowercased and stripped of extra whitespace
    """
    return str(value).lower().strip()


def color_negative_red(val):
    """
    Takes a scalar and returns a string with
    the css property `'color: red'` for ratios
    greater than or equal to 0.7, black otherwise.
    """
    color = "black"
    try:
        color = "red" if float(val) >= RATIO_THRESHHOLD and float(val) <= 1 else "black"
    except ValueError:
        print("Not a float")
    return "color: %s" % color


# Prematch by comparing lower cased and trimmed values
for i, landlord in landlords.iterrows():
    landlord_domain = ""
    common_domains = ["gmail.com", "yahoo.com", "NaN", "nan", "aol.com"]
    if str("@") in str(landlord["Landlord Email"]):
        landlord_domain = lower_strip(landlord["Landlord Email"]).split("@")[1]

    for ii, tenant in filtered_tenants.iterrows():
        tenant_domain = ""
        if str("@") in str(tenant["Landlord Email"]):
            tenant_domain = lower_strip(tenant["Landlord Email"]).split("@")[1]

        if lower_strip(tenant["Tenant Email"]) == lower_strip(landlord["Tenant Email"]):
            matched_list.append(
                {"tenant": tenant, "landlord": landlord, "match_type": "Email"}
            )
            continue
        elif lower_strip(tenant["Requested for"]) == lower_strip(
            landlord["Requested for"]
        ):
            matched_list.append(
                {"tenant": tenant, "landlord": landlord, "match_type": "requested for"}
            )
            continue
        elif lower_strip(tenant["Address line 1"]) == lower_strip(
            landlord["Address line 1"]
        ):
            matched_list.append(
                {"tenant": tenant, "landlord": landlord, "match_type": "address line 1"}
            )
            continue
        elif lower_strip(tenant["Landlord Name"]) == lower_strip(
            landlord["Landlord Name"]
        ):
            matched_list.append(
                {"tenant": tenant, "landlord": landlord, "match_type": "landlord name"}
            )
            continue
        elif lower_strip(tenant["Landlord Email"]) == lower_strip(
            landlord["Landlord Email"]
        ):
            matched_list.append(
                {"tenant": tenant, "landlord": landlord, "match_type": "landlord email"}
            )
            continue
        elif (
            landlord_domain not in common_domains
            and landlord_domain != ""
            and tenant_domain == landlord_domain
        ):
            matched_list.append(
                {"tenant": tenant, "landlord": landlord, "match_type": "domain"}
            )
            continue

# Get Levenshtein Ratio for each match
for match in matched_list:
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
        lower_strip(match["tenant"]["Zip Code"]),
        lower_strip(match["landlord"]["Zip Code"]),
    )
    ratios = {
        "Tenant RITM": match["tenant"]["Number"],
        "Landlord RITM": match["landlord"]["Number"],
        "match_type": match["match_type"],
        #
        "tenant_landlord_email": match["tenant"]["Landlord Email"],
        "landlord_email": match["landlord"]["Landlord Email"],
        "Landlord Email ratio": landlord_email,
        #
        "tenant_email": match["tenant"]["Tenant Email"],
        "landlord_tenant_email": match["landlord"]["Tenant Email"],
        "Tenant Email ratio": tenant_email,
        #
        "Tenant address": f"{lower_strip(match['tenant']['Address line 1'])} {lower_strip(match['tenant']['Address line 2'])}",
        "Landlord Address": f"{lower_strip(match['landlord']['Address line 1'])} {lower_strip(match['landlord']['Address line 2'])}",
        "Address line 1_2 ratio": address_line,
        #
        "Tenant requested for": lower_strip(match["tenant"]["Requested for"]),
        "Landlord requested for": lower_strip(match["landlord"]["Requested for"]),
        "Requested for ratio": requested_for,
        #
        "tenant landlord name": lower_strip(match["tenant"]["Landlord Name"]),
        "landlord name": lower_strip(match["landlord"]["Landlord Name"]),
        "Landlord Name ratio": landlord_name,
        #
        "tenant zip code": lower_strip(match["tenant"]["Zip Code"]),
        "landlord zip code": lower_strip(match["landlord"]["Zip Code"]),
        "zip code ratio": zip_code,
        #
        "average": (
            landlord_email
            + landlord_name
            + tenant_email
            + address_line
            + requested_for
            + zip_code
        )
        / 6,
    }
    results.append(ratios)

print(f"{len(results)} potential matches found")

results_df = pd.DataFrame(results)
results_df = results_df.style.applymap(color_negative_red)
results_df.to_excel(OUTPUT_FILENAME)

print("All done")
