# Levenshtein String Comparison

## Requirements

1. Python 3.9+

## Usage

1. git clone
2. cd ..
3. pip install -r requirements.txt
4. Export the [HP Unmatched report](https://dcerapprod.servicenowservices.com/sys_report_template.do?jvar_report_id=d8c6ed561b1c74109704dd39bc4bcb72) and save it to the root of this folders as `landlords.xlsx`
5. Export the [Tenant Unmatched report](https://dcerapprod.servicenowservices.com/sys_report_template.do?jvar_report_id=8241d2e41b6434509704dd39bc4bcb4c) and save it to the root of this folders as `tenants.xlsx`
6. run `python ritm_matcher.py`.
7. Once the script finishes running verify the new `matches.xlxs` file has been created
8. Zip and encrypt the file output and send to Linus, Kara and Ashley.

### Instance

- The above process is for DC but it should work exactly the same for North Dakota and Nebraska.
  - You only need to recreate the 2 reports in that instance.
