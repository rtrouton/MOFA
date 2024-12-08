import os
import requests
import xml.etree.ElementTree as ET
from hashlib import sha256, sha1
from xml.dom import minidom
import json
from datetime import datetime
import time
import pytz
import re
import yaml

def get_current_date_time():
    # Get current UTC time and convert it to Eastern Time (or any other timezone)
    utc_now = datetime.now(pytz.utc)  # Get current UTC time with tz info
    eastern_time = utc_now.astimezone(pytz.timezone('US/Eastern'))  # Convert to Eastern Time

    # Format the date and time as needed (e.g., 12/06/2024 04:30 PM Eastern)
    formatted_date_time = eastern_time.strftime('%B %d, %Y %I:%M %p %Z')  # 'December 06, 2024 04:30 PM Eastern'

    return formatted_date_time

# Call the function to test it
last_update_date_time = get_current_date_time()
print(last_update_date_time)

# Define app-specific configurations
apps = {
    "Microsoft Office Suite": {
        "url": "https://officecdnmac.microsoft.com/pr/C1297A47-86C4-4C1F-97FA-950631F94777/MacAutoupdate/0409XCEL2019.xml",
        "manual_entries": {
            "CFBundleVersion": "com.microsoft.office",
            "latest_download": "https://go.microsoft.com/fwlink/?linkid=525133",
            "full_version": "https://go.microsoft.com/fwlink/?linkid=525133",
            "application_id": "Data sourced from Excel (not manually specified)",
            "application_name": "Data sourced from Excel (not manually specified)",
        },
        "keys": {
            "short_version": "Title",
            "full_version": "Update Version",
            "update_download": "NA",
            "last_updated": "Date",
            "min_os": "Minimum OS"
        }
    },
    "Microsoft BusinessPro Suite": {
        "url": "https://officecdnmac.microsoft.com/pr/C1297A47-86C4-4C1F-97FA-950631F94777/MacAutoupdate/0409XCEL2019.xml",
        "manual_entries": {
            "CFBundleVersion": "com.microsoft.office",
            "latest_download": "https://go.microsoft.com/fwlink/?linkid=2009112",
            "full_version": "https://go.microsoft.com/fwlink/?linkid=2009112",
            "application_id": "Data sourced from Excel (not manually specified)",
            "application_name": "Data sourced from Excel (not manually specified)",
        },
        "keys": {
           "short_version": "Title",
            "full_version": "Update Version",
            "update_download": "NA",
            "last_updated": "Date",
            "min_os": "Minimum OS"
        }
    },    
    "Word": {
        "url": "https://officecdnmac.microsoft.com/pr/C1297A47-86C4-4C1F-97FA-950631F94777/MacAutoupdate/0409MSWD2019.xml",
        "manual_entries": {
            "CFBundleVersion": "com.microsoft.word",
            "latest_download": "https://go.microsoft.com/fwlink/?linkid=525134",
        },
        "keys": {
            "application_id": "Application ID",
            "application_name": "Application Name",
            "short_version": "Title",
            "full_version": "Update Version",
            "update_download": "FullUpdaterLocation",
            "last_updated": "Date",
            "min_os": "Minimum OS"
        }
    },
    "Excel": {
        "url": "https://officecdnmac.microsoft.com/pr/C1297A47-86C4-4C1F-97FA-950631F94777/MacAutoupdate/0409XCEL2019.xml",
        "manual_entries": {
            "CFBundleVersion": "com.microsoft.excel",
            "latest_download": "https://go.microsoft.com/fwlink/?linkid=525135",
        },
        "keys": {
            "application_id": "Application ID",
            "application_name": "Application Name",
            "short_version": "Title",
            "full_version": "Update Version",
            "update_download": "FullUpdaterLocation",
            "last_updated": "Date",
            "min_os": "Minimum OS"
        }
    },
    "PowerPoint": {
        "url": "https://officecdnmac.microsoft.com/pr/C1297A47-86C4-4C1F-97FA-950631F94777/MacAutoupdate/0409PPT32019.xml",
        "manual_entries": {
            "CFBundleVersion": "com.microsoft.powerpoint",
            "latest_download": "https://go.microsoft.com/fwlink/?linkid=525136",
        },
        "keys": {
            "application_id": "Application ID",
            "application_name": "Application Name",
            "short_version": "Title",
            "full_version": "Update Version",
            "update_download": "FullUpdaterLocation",
            "last_updated": "Date",
            "min_os": "Minimum OS"
        }
    },
    "Outlook": {
        "url": "https://officecdnmac.microsoft.com/pr/C1297A47-86C4-4C1F-97FA-950631F94777/MacAutoupdate/0409OPIM2019.xml",
        "manual_entries": {
            "CFBundleVersion": "com.microsoft.outlook",
            "latest_download": "https://go.microsoft.com/fwlink/?linkid=2228621",
        },
        "keys": {
            "application_id": "Application ID",
            "application_name": "Application Name",
            "short_version": "Title",
            "full_version": "Update Version",
            "update_download": "FullUpdaterLocation",
            "last_updated": "Date",
            "min_os": "Minimum OS"
        }
    },
    "OneNote": {
        "url": "https://officecdnmac.microsoft.com/pr/C1297A47-86C4-4C1F-97FA-950631F94777/MacAutoupdate/0409ONMC2019.xml",
        "manual_entries": {
            "CFBundleVersion": "com.microsoft.onenote",
            "latest_download": "https://go.microsoft.com/fwlink/?linkid=820886",
        },
        "keys": {
            "application_id": "Application ID",
            "application_name": "Application Name",
            "short_version": "Title",
            "full_version": "Update Version",
            "update_download": "FullUpdaterLocation",
            "last_updated": "Date",
            "min_os": "Minimum OS"
        }
    },
    "OneDrive": {
        "url": "https://officecdnmac.microsoft.com/pr/C1297A47-86C4-4C1F-97FA-950631F94777/MacAutoupdate/0409ONDR18.xml",
        "manual_entries": {
            "CFBundleVersion": "com.microsoft.onedrive",
            "latest_download": "https://go.microsoft.com/fwlink/?linkid=823060",
            "application_name": "OneDrive.app",
        },
        "keys": {
            "application_id": "Application ID",
            "short_version": "Title",
            "full_version": "Update Version",
            "update_download": "Location",
            "last_updated": "Date",
            "min_os": "Minimum OS"
        }
    },
    "Skype": {
        "url": "https://officecdnmac.microsoft.com/pr/C1297A47-86C4-4C1F-97FA-950631F94777/MacAutoupdate/0409MSFB16.xml",
        "manual_entries": {
            "CFBundleVersion": "com.microsoft.skypeforbusiness",
            "latest_download": "https://go.microsoft.com/fwlink/?linkid=832978",
            "application_name": "Skype for Business.app",
        },
        "keys": {
            "application_id": "Application ID",
            "short_version": "Title",
            "full_version": "Update Version",
            "update_download": "Location",
            "last_updated": "Date",
            "min_os": "Minimum OS"
        }
    },
    "Teams": {
        "url": "https://statics.teams.cdn.office.net/production-osx/24295.606.3238.6194/0409TEAMS21.xml",
        "manual_entries": {
            "CFBundleVersion": "com.microsoft.teams",
            "latest_download": "https://go.microsoft.com/fwlink/?linkid=2249065",
            "application_name": "Microsoft Teams.app",
        },
        "keys": {
            "application_id": "Application ID",
            "short_version": "Title",
            "full_version": "Update Version",
            "update_download": "Location",
            "last_updated": "Date",
            "min_os": "Minimum OS"
        }
    },    "Intune": {
        "url": "https://officecdnmac.microsoft.com/pr/C1297A47-86C4-4C1F-97FA-950631F94777/MacAutoupdate/0409IMCP01.xml",
        "manual_entries": {
            "CFBundleVersion": "com.microsoft.intune.companyportal",
            "latest_download": "https://go.microsoft.com/fwlink/?linkid=853070",
            "application_name": "Company Portal.app",
        },
        "keys": {
            "application_id": "Application ID",
            "short_version": "Title",
            "full_version": "Update Version",
            "update_download": "Location",
            "last_updated": "Date",
            "min_os": "Minimum OS Version"
        }
    },
    "Edge": {
        "url": "https://officecdnmac.microsoft.com/pr/C1297A47-86C4-4C1F-97FA-950631F94777/MacAutoupdate/0409EDGE01.xml",
        "manual_entries": {
            "CFBundleVersion": "com.microsoft.edgemac",
            "latest_download": "https://go.microsoft.com/fwlink/?linkid=2093504",
            "application_name": "Microsoft Edge.app",
        },
        "keys": {
            "application_id": "Application ID",
            "short_version": "Title",
            "full_version": "Update Version",
            "update_download": "Location",
            "last_updated": "Date",
            "min_os": "Minimum OS"
        }
    },
    "Defender For Endpoint": {
        "url": "https://officecdnmac.microsoft.com/pr/C1297A47-86C4-4C1F-97FA-950631F94777/MacAutoupdate/0409WDAV00.xml",
        "manual_entries": {
            "CFBundleVersion": "com.microsoft.defender.endpoint",
            "latest_download": "https://go.microsoft.com/fwlink/?linkid=2097502",
            "application_name": "Microsoft Defender.app",
        },
        "keys": {
            "application_id": "Application ID",
            "short_version": "Title",
            "full_version": "Update Version",
            "update_download": "Location",
            "last_updated": "Date",
            "min_os": "Minimum OS"
        }
    },
    "Defender for Consumers": {
        "url": "https://officecdnmac.microsoft.com/pr/C1297A47-86C4-4C1F-97FA-950631F94777/MacAutoupdate/0409WDAVCONSUMER.xml",
        "manual_entries": {
            "CFBundleVersion": "com.microsoft.defender.endpoint",
            "latest_download": "https://go.microsoft.com/fwlink/?linkid=2247001",
            "application_name": "Microsoft Defender.app",
        },
        "keys": {
            "application_id": "Application ID",
            "short_version": "Title",
            "full_version": "Update Version",
            "update_download": "Location",
            "last_updated": "Date",
            "min_os": "Minimum OS"
        }
    },
    "Defender Shim": {
        "url": "https://officecdnmac.microsoft.com/pr/C1297A47-86C4-4C1F-97FA-950631F94777/MacAutoupdate/0409WDAVSHIM.xml",
        "manual_entries": {
            "CFBundleVersion": "com.microsoft.defender.endpoint",
            "application_name": "N/A",
        },
        "keys": {
            "application_id": "Application ID",
            "latest_download": "Location",
            "short_version": "Title",
            "full_version": "Update Version",
            "update_download": "Location",
            "last_updated": "Date",
            "min_os": "Minimum OS"
        }
    },
    "Windows App": {
        "url": "https://officecdnmac.microsoft.com/pr/C1297A47-86C4-4C1F-97FA-950631F94777/MacAutoupdate/0409MSRD10.xml",
        "manual_entries": {
            "CFBundleVersion": "com.microsoft.windows.app",
            "latest_download": "https://go.microsoft.com/fwlink/?linkid=868963",
            "application_name": "Windows App.app",
        },
        "keys": {
            "application_id": "Application ID",
            "short_version": "Title",
            "full_version": "Update Version",
            "update_download": "Location",
            "last_updated": "Date",
            "min_os": "Minimum OS Version"
        }
    },
    "Visual": {
        "url": "https://update.code.visualstudio.com/api/update/darwin-universal/stable/384ff7382de624fb94dbaf6da11977bba1ecd427",
        "manual_entries": {
            "CFBundleVersion": "com.microsoft.visualstudio",
            "latest_download": "https://go.microsoft.com/fwlink/?linkid=2156837",
            "application_name": "Visual Studio Code.app",
        },
        "keys": {
            "application_id": "Application ID",
            "short_version": "name",
            "full_version": "productVersion",
            "update_download": "url",
            "last_updated": "timestamp",
            "min_os": "Minimum OS"
        }
    },
    "MAU": {
        "url": "https://officecdnmac.microsoft.com/pr/C1297A47-86C4-4C1F-97FA-950631F94777/MacAutoupdate/0409MSau04.xml",
        "manual_entries": {
            "CFBundleVersion": "com.microsoft.autoupdate",
            "latest_download": "https://go.microsoft.com/fwlink/?linkid=830196",
            "application_name": "Microsoft AutoUpdate.app",
        },
        "keys": {
            "application_id": "Application ID",
            "short_version": "Title",
            "full_version": "Update Version",
            "update_download": "Location",
            "last_updated": "Date",
            "min_os": "Minimum OS"
        }
    },
    "Licensing Helper Tool": {
        "url": "https://officecdnmac.microsoft.com/pr/C1297A47-86C4-4C1F-97FA-950631F94777/MacAutoupdate/0409OLIC02.xml",
        "manual_entries": {
            "CFBundleVersion": "com.microsoft.office.licensingV2.helper",
            "latest_download": "https://go.microsoft.com/fwlink/p/?linkid=2181269",
            "application_name": "N/A",
        },
        "keys": {
            "application_id": "Application ID",
            "short_version": "Title",
            "full_version": "Update Version",
            "update_download": "Location",
            "last_updated": "Date",
            "min_os": "Minimum OS"
        }
    },
    "Quick Assist": {
        "url": "https://officecdnmac.microsoft.com/pr/C1297A47-86C4-4C1F-97FA-950631F94777/MacAutoupdate/0409MSQA01.xml",
        "manual_entries": {
            "CFBundleVersion": "com.microsoft.quickassist",
            "latest_download": "https://go.microsoft.com/fwlink/?linkid=2181269",
        },
        "keys": {
            "application_id": "Application ID",
            "application_name": "Application Name",
            "latest_download": "Location",
            "short_version": "Update Version",
            "full_version": "Update Version",
            "update_download": "Location",
            "last_updated": "Date",
            "min_os": "Minimum OS"
        }
    },
    "Remote Help": {
        "url": "https://officecdnmac.microsoft.com/pr/C1297A47-86C4-4C1F-97FA-950631F94777/MacAutoupdate/0409MSRH01.xml",
        "manual_entries": {
            "CFBundleVersion": "com.microsoft.remotehelp",
            "latest_download": "https://go.microsoft.com/fwlink/?linkid=2181269",
        },
        "keys": {
            "application_id": "Application ID",
            "application_name": "Application Name",
            "latest_download": "Location",
            "short_version": "Update Version",
            "full_version": "Update Version",
            "update_download": "Location",
            "last_updated": "Date",
            "min_os": "Minimum OS"
        }
    }
}

# Capture the current last update date and time
last_update_date_time = get_current_date_time()

# Initialize root element for combined XML
root = ET.Element("latest")

# Add the last update date and time element to the XML
last_update_element = ET.SubElement(root, "last_updated")
last_update_element.text = last_update_date_time  # Value from get_current_date_time()

# Function to read existing XML data from latest.xml
def read_existing_xml(filename):
    if not os.path.exists(filename):
        return {}

    try:
        tree = ET.parse(filename)
        root = tree.getroot()
        existing_data = {}
        for package in root.findall("package"):
            app_name = package.find("name").text
            last_updated = package.find("last_updated").text
            existing_data[app_name] = {
                "last_updated": last_updated,
                "data": {child.tag: child.text for child in package}
            }
        return existing_data
    except Exception as e:
        print(f"Error reading existing XML: {e}")
        return {}

# Read existing data from latest.xml
existing_data = read_existing_xml("latest.xml")

# Function to fetch and process an app's data (either XML or JSON)
def fetch_and_process(app_name, config):
    try:
        print(f"Processing {app_name}...")
        response = requests.get(config["url"])
        response.raise_for_status()

        # Check if the response is in JSON format
        if response.headers['Content-Type'].startswith('application/json'):
            app_data = response.json()
            extracted_data = process_json_data(app_data, config)
        else:
            app_root = ET.fromstring(response.content)
            app_data = app_root.find(".//dict")
            extracted_data = process_xml_data(app_data, config)

        # Add manual entries
        extracted_data.update(config["manual_entries"])

        # Check if last_updated has changed or if sha1 or sha256 are "N/A"
        if app_name in existing_data:
            existing_last_updated = existing_data[app_name]["last_updated"]
            existing_sha1 = existing_data[app_name]["data"].get("sha1", "N/A")
            existing_sha256 = existing_data[app_name]["data"].get("sha256", "N/A")
            extracted_last_updated = extracted_data.get("last_updated", "N/A")
            extracted_sha1 = extracted_data.get("sha1", "N/A")
            extracted_sha256 = extracted_data.get("sha256", "N/A")

            # If any of the required fields are "N/A" or if last_updated differs
            if (existing_last_updated == extracted_last_updated and
                existing_sha1 != "N/A" and existing_sha256 != "N/A"):
                print(f"No update for {app_name}.")
                add_to_combined_xml(app_name, existing_data[app_name]["data"])
            else:
                print(f"Update detected for {app_name}. Computing SHA...")
                # Compute SHA only for updated apps or apps with "N/A" fields
                download_url = extracted_data.get("latest_download")
                extracted_data["sha1"] = compute_sha1(download_url) if download_url else "N/A" ## COMMENT THESE OUT FOR QUICK TESTING
                extracted_data["sha256"] = compute_sha256(download_url) if download_url else "N/A" ## COMMENT THESE OUT FOR QUICK TESTING
                add_to_combined_xml(app_name, extracted_data)
        else:
            print(f"New app {app_name} detected. Computing SHA...")
            download_url = extracted_data.get("latest_download")
            extracted_data["sha1"] = compute_sha1(download_url) if download_url else "N/A" ## COMMENT THESE OUT FOR QUICK TESTING
            extracted_data["sha256"] = compute_sha256(download_url) if download_url else "N/A" ## COMMENT THESE OUT FOR QUICK TESTING
            add_to_combined_xml(app_name, extracted_data)

    except Exception as e:
        print(f"Error processing {app_name}: {e}")
        # Use existing data if processing fails
        if app_name in existing_data:
            print(f"Reverting to existing data for {app_name}.")
            add_to_combined_xml(app_name, existing_data[app_name]["data"])

# Function to process XML data
def process_xml_data(app_data, config):
    extracted_data = {}
    for field, key in config["keys"].items():
        extracted_data[field] = find_key_value(app_data, key)

    last_updated = extracted_data.get("last_updated", "N/A")
    extracted_data["last_updated"] = convert_last_updated(last_updated)

    if 'short_version' in extracted_data:
        extracted_data["short_version"] = re.sub(r'[a-zA-Z]', '', extracted_data["short_version"]).lstrip()

    return extracted_data

# Function to process JSON data
def process_json_data(app_data, config):
    extracted_data = {}
    for field, key in config["keys"].items():
        extracted_data[field] = app_data.get(key, "N/A")

    last_updated = extracted_data.get("last_updated", "N/A")
    extracted_data["last_updated"] = convert_last_updated(last_updated)

    if 'short_version' in extracted_data:
        extracted_data["short_version"] = re.sub(r'[a-zA-Z]', '', extracted_data["short_version"]).lstrip()

    return extracted_data

# Helper function to find a key's value in the XML
def find_key_value(element, key_name):
    found = False
    for child in element:
        if found:
            return child.text if child.text else "N/A"
        if child.tag == "key" and child.text == key_name:
            found = True
    return "N/A"

# Function to compute SHA1 hash
def compute_sha1(url):
    try:
        # Use allow_redirects=True to follow redirects
        response = requests.get(url, stream=True, allow_redirects=True)
        response.raise_for_status()  # Raise exception for HTTP errors
        hasher = sha1()
        for chunk in response.iter_content(chunk_size=8192):
            hasher.update(chunk)
        return hasher.hexdigest()
    except Exception as e:
        print(f"Error computing SHA1 for {url}: {e}")
        return "N/A"

# Function to compute SHA256 hash
def compute_sha256(url):
    try:
        # Use allow_redirects=True to follow redirects
        response = requests.get(url, stream=True, allow_redirects=True)
        response.raise_for_status()  # Raise exception for HTTP errors
        hasher = sha256()
        for chunk in response.iter_content(chunk_size=8192):
            hasher.update(chunk)
        return hasher.hexdigest()
    except Exception as e:
        print(f"Error computing SHA256 for {url}: {e}")
        return "N/A"

def add_to_combined_xml(app_name, data):
    package = ET.SubElement(root, "package")

    # Add the elements in the specified order
    order = [
        "name",
        "application_id",
        "application_name",
        "CFBundleVersion",
        "short_version",
        "full_version",
        "last_updated",
        "min_os",
        "update_download",
        "latest_download",
        "sha1",
        "sha256",
    ]

    for key in order:
        if key == "name":  # Ensure the name element is added only once
            name_element = ET.SubElement(package, key)
            name_element.text = app_name
        elif key in data:
            ET.SubElement(package, key).text = data[key]


# Function to convert last_updated to human-readable date
def convert_last_updated(last_updated):
    if isinstance(last_updated, int) or (last_updated.isdigit() if isinstance(last_updated, str) else False):
        return time.strftime('%B %d, %Y', time.gmtime(int(last_updated) / 1000))
    elif isinstance(last_updated, str):
        try:
            date_obj = datetime.fromisoformat(last_updated.replace('Z', '+00:00'))
            return date_obj.strftime('%B %d, %Y')
        except ValueError:
            pass
    return "N/A"

# Process each app and populate combined XML
for app_name, config in apps.items():
    fetch_and_process(app_name, config)

# Pretty print the XML
def pretty_print_xml(element):
    rough_string = ET.tostring(element, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="    ")

# Save the updated XML
output_file = "latest.xml"
pretty_xml = pretty_print_xml(root)

with open(output_file, "w", encoding="utf-8") as f:
    f.write(pretty_xml)

print(f"Combined and prettified XML generated: {output_file}")

# Generate and save YAML output
yaml_data = {
    "last_updated": last_update_date_time,
    "packages": [
        {"name": app_name, "data": app_data["data"]}
        for app_name, app_data in existing_data.items()
    ],
}

yaml_output_file = "latest.yaml"
with open(yaml_output_file, "w", encoding="utf-8") as yaml_file:
    yaml.dump(yaml_data, yaml_file, default_flow_style=False)

print(f"YAML output generated: {yaml_output_file}")
