#!/usr/bin/env python
# coding: utf-8

import os
import datetime

import Levenshtein as lev
import pandas as pd


# print(start_time)
THRESHOLD = 7.5
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

landlords = pd.read_excel(LANDLORDS_PATH)
# landlords.fillna("Empty")
print(f"{len(landlords)} Unmatched Landlord RITMs")

tenants = pd.read_excel(TENANTS_PATH)
# tenants.fillna("Empty")
print(f"{len(tenants)} Unmatched Tenant RITMs")

# tenants = tenants[tenants.Number.isin(UNMATCHED_RITMS)]
print(f"Checking {len(tenants)} tenants for matches.")
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
    greater than or equal to THRESHOLD / 10,
    black otherwise.
    """
    color = "black"
    if isinstance(val, datetime.datetime):
        pass
    else:
        try:
            color = (
                "red" if float(val) >= THRESHOLD / 10 and float(val) <= 1 else "black"
            )
        except ValueError:
            color = "black"
    return "color: %s" % color


def sanitize_data(values, ratio):
    """
    returns the lev ratio unless one or more of the
    parameters passed isn't a string,
    then it returns 0
    """
    any_blanks = False
    for i in values:
        if not isinstance(i, str) or i == "" or i == "nan":
            any_blanks = True
    return 0 if any_blanks else ratio


def remove_nan(value):
    if not isinstance(value, str) or value == "nan":
        return ""
    else:
        return value


for i, landlord in landlords.iterrows():
    """
    Prematch by comparing lower cased and
    trimmed values.
    """
    landlord_domain = ""
    landlord_tenant_domain = ""  #
    print(f"Checking landlord #{i + 1}")
    if str("@") in str(landlord["Landlord Email"]):
        landlord_domain = lower_strip(landlord["Landlord Email"]).split("@")[1]
    if str("@") in str(landlord["Tenant Email"]):
        landlord_tenant_domain = lower_strip(landlord["Tenant Email"]).split("@")[1]
        #
        #
    for ii, tenant in tenants.iterrows():
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
        elif (
            lower_strip(tenant["Requested for"])
            == f"{lower_strip(landlord['Tenant first name'])} {lower_strip(landlord['Tenant last name'])}"
        ):
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
    tenant_name = lev.ratio(
        remove_nan(lower_strip(match["tenant"]["Requested for"])),
        f"{remove_nan(lower_strip(match['landlord']['Tenant first name']))} {remove_nan(lower_strip(match['landlord']['Tenant last name']))}",
    )
    landlord_name = lev.ratio(
        lower_strip(match["tenant"]["Landlord Name"]),
        lower_strip(match["landlord"]["Landlord Name"]),
    )
    landlord_name = sanitize_data(
        [
            lower_strip(match["tenant"]["Landlord Name"]),
            lower_strip(match["landlord"]["Landlord Name"]),
        ],
        landlord_name,
    )
    landlord_email = lev.ratio(
        lower_strip(match["tenant"]["Landlord Email"]),
        lower_strip(match["landlord"]["Landlord Email"]),
    )
    ratios = {
        "Tenant RITM": match["tenant"]["Number"],
        "LL RITM": match["landlord"]["Number"],
        #
        "Tenant Multiple Matches": "Not sure yet",
        "LL Multiple Matches": "Not sure yet",
        "Match Type": match["match_type"],
        #
        "Tenant Name (Tenant)": remove_nan(
            lower_strip(match["tenant"]["Requested for"])
        ),
        "Tenant Name (LL)": f"{remove_nan(lower_strip(match['landlord']['Tenant first name']))} {remove_nan(lower_strip(match['landlord']['Tenant last name']))}",
        "Tenant Name Comparison": tenant_name,
        #
        "Tenant Address 1 + 2 (Tenant)": f"{lower_strip(match['tenant']['Address line 1'])} {remove_nan(lower_strip(match['tenant']['Address line 2']))}",
        "Tenant Address 1 + 2 (LL)": f"{lower_strip(match['landlord']['Address line 1'])} {remove_nan(lower_strip(match['landlord']['Address line 2']))}",
        "Address Line Comparison": address_line,
        #
        "Tenant Email (Tenant)": match["tenant"]["Tenant Email"],
        "Tenant Email (LL)": match["landlord"]["Tenant Email"],
        "Tenant Email Comparison": tenant_email,
        #
        "Landlord Name (Tenant)": remove_nan(match["tenant"]["Landlord Name"]),
        "Landlord Name (LL)": remove_nan(match["landlord"]["Landlord Name"]),
        "Landlord Name Comparison": landlord_name,
        #
        "Landlord Email (Tenant)": match["tenant"]["Landlord Email"],
        "Landlord Email (LL)": match["landlord"]["Landlord Email"],
        "Landlord Email Comparison": sanitize_data(
            [match["tenant"]["Landlord Email"], match["landlord"]["Landlord Email"]],
            landlord_email,
        ),
        #
        "Match Score": float(
            "{:.2f}".format(
                (tenant_email * 3)
                + (tenant_name * 2.5)
                + (address_line * 1.5)
                + (landlord_name * 1.5)
                + (landlord_email * 1.5)
            )
        ),
        # Created
        "Created": match["tenant"]["Created"],
    }
    results.append(ratios)


results_df = pd.DataFrame(results)
print(f"{len(results_df)} potential matches found before filtering.")
unfiltered = results_df
results_df = results_df.loc[(results_df["Match Score"]) >= 7.5]
# ritm_list = results_df[""].tolist()

results_df["Tenant Multiple Matches"] = results_df.duplicated(
    subset=["Tenant RITM"], keep=False
)
results_df["LL Multiple Matches"] = results_df.duplicated(
    subset=["LL RITM"], keep=False
)
unfiltered["Tenant Multiple Matches"] = unfiltered.duplicated(
    subset=["Tenant RITM"], keep=False
)
unfiltered["LL Multiple Matches"] = unfiltered.duplicated(
    subset=["LL RITM"], keep=False
)

print(f"{len(results_df)} potential matches found.")

results_df = results_df.sort_values(
    by=["Tenant RITM", "Match Score"], ascending=[True, False]
)
unfiltered = unfiltered.sort_values(
    by=["Tenant RITM", "Match Score"], ascending=[True, False]
)

results_df = results_df.style.applymap(color_red)
with pd.ExcelWriter(OUTPUT_FILENAME) as writer:
    results_df.to_excel(writer, sheet_name="Filtered", index=False)

print(f"{OUTPUT_FILENAME} file created at {CURRENT_DIRECTORY}.")
print("All done! ♪┏(°.°)┛┗(°.°)┓┗(°.°)┛┏(°.°)┓ ♪")
# print(time.clock() - start_time, "seconds")

# TODO Make the output filename a command line argument.
# TODO run again with fillna done correctly
