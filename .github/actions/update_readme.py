import xml.etree.ElementTree as ET
from datetime import datetime
import pytz
import logging

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

def parse_latest_xml(file_path):
    logging.info(f"Parsing XML file: {file_path}")
    
    # Parse the XML file
    tree = ET.parse(file_path)
    root = tree.getroot()
    logging.debug("XML file parsed successfully")

    # Extract global last_updated
    global_last_updated = root.find("last_updated").text.strip()
    logging.info(f"XML last_updated: {global_last_updated}")

    # Extract package details
    packages = {}
    for package in root.findall("package"):
        name = package.find("name").text.strip().lower()  # Store by lowercase name for easy access
        packages[name] = {
            "name": name,
            "application_id": package.find("application_id").text.strip(),
            "application_name": package.find("application_name").text.strip(),
            "short_version": package.find("short_version").text.strip(),
            "full_version": package.find("full_version").text.strip(),
            "min_os": package.find("min_os").text.strip(),
            "update_download": package.find("update_download").text.strip(),
            "latest_download": package.find("latest_download").text.strip(),
            "sha256": package.find("sha256").text.strip(),
        }
        # logging.debug(f"Extracted package: {packages[name]}") # Uncomment to see all package details
    
    return global_last_updated, packages

def generate_readme_content(global_last_updated, packages):
    logging.info("Generating README content")
    
    # Set timezone to US/Eastern (EST/EDT)
    eastern = pytz.timezone('US/Eastern')
    
    # Get the current time in UTC and convert to EST
    current_time = datetime.now(pytz.utc).astimezone(eastern).strftime("%B %d, %Y %I:%M %p %Z")
    logging.debug(f"Current time (EST): {current_time}")
    
    content = f"""# **MOFA**
**M**icrosoft **O**ffice **F**eed for **A**pple

Welcome to the **MOFA** repository! This centralized and up-to-date resource offers Microsoft Office downloads for macOS, data feeds, along with comprehensive support documentation for Mac administrators. It aims to replace the now-defunct [**MacAdmins.software**](https://macadmins.software), and we would like to extend a special thanks to [**Paul Bowden**](https://github.com/pbowden-msft) for his exceptional contributions to the Mac Admins community.

We encourage active community participation—feel free to fork the repository, ask questions, or contribute additional insights to ensure its ongoing accuracy and usefulness. Together, we can keep this resource thriving for everyone.

### 🌟 **Current Focus Areas**  

- **✅ Feature/Improvement:** Automated `latest.xml` (Checks every 4 hours with Microsoft sources) – *Completed!*  
- **✅ Feature/Improvement:** Automating workflow to update `README.md` – *Completed!* 
- **⏸️ Feature/Improvement:** Adding more links/documentation – *Paused* (Community support needed! 🆘)
- **🔄 Feature/Improvement:**  Add iOS/MacOS AppStore Info – *In Progress*
- **⏸️ Future:** Add JSON Output – *Paused* (Waiting on iOS/MacOS AppStore Info)
- **🚀 Future:** GitHub webpage with custom domain – *Waiting on funding/support 💸*

I’m not here to make a profit—just here to ensure knowledge remains open and free for all, check the button below to support MOFA:

<a href="https://www.buymeacoffee.com/cocopuff2u"><img src="https://img.buymeacoffee.com/button-api/?text=Support This Project&emoji=💻&slug=cocopuff2u&button_colour=FFDD00&font_colour=000000&font_family=Bree&outline_colour=000000&coffee_colour=ffffff" style="height: 40px;" /></a>  

## <img src="images/Microsoft_Logo_512px.png" alt="Download Image" width="20"></a> Microsoft Office Installers

<sup>All links below direct to Microsoft's official Content Delivery Network (CDN).</sup>
<sup>The links provided will always download the latest version offered by Microsoft. However, the version information listed below reflects the version available at the time of this update.</sup>

<sup>_Last Updated: <code style="color : mediumseagreen">{global_last_updated}</code> [**_Raw XML_**](/latest.xml) [**_Raw YAML_**](/latest.yaml) (Automatically Updated every 4 hours)_</sup>

| **Product Package** | **CFBundle Version** | **CFBundle Identifier** | **Download** |
|----------------------|----------------------|--------------------------|--------------|
| **Microsoft** <sup>365/2021/2024</sup> **Office Suite Installer**<br><a href="https://learn.microsoft.com/en-us/officeupdates/release-notes-office-for-mac" style="text-decoration: none;"><small>_Release Notes_</small></a><br><sub>_(Includes Word, Excel, PowerPoint, Outlook, OneNote, OneDrive, and MAU)_</sub> | `{get_package_detail(packages, 'Microsoft Office Suite', 'short_version')}` | com.microsoft.office | <a href="https://go.microsoft.com/fwlink/?linkid=525133"><img src="images/suite.png" alt="Download Image" width="80"></a> |
| **Microsoft** <sup>365/2021/2024</sup> **BusinessPro Suite Installer**<br><sub>_(Includes Word, Excel, PowerPoint, Outlook, OneNote, OneDrive, Teams, Defender Shim, and MAU)_</sub> | `{get_package_detail(packages, 'Microsoft BusinessPro Suite', 'short_version')}` | com.microsoft.office | <a href="https://go.microsoft.com/fwlink/?linkid=2009112"><img src="images/suite.png" alt="Download Image" width="80"></a> |
| **Word** <sup>365/2021/2024</sup> **</sup> Standalone Installer** | `{get_package_detail(packages, 'Word', 'short_version')}` | com.microsoft.word | <a href="https://go.microsoft.com/fwlink/?linkid=525134"><img src="images/MSWD_512x512x32.png" alt="Download Image" width="80"></a> |
| **Excel** <sup>365/2021/2024</sup> **Standalone Installer** | `{get_package_detail(packages, 'Excel', 'short_version')}` | com.microsoft.excel | <a href="https://go.microsoft.com/fwlink/?linkid=525135"><img src="images/XCEL_512x512x32.png" alt="Download Image" width="80"></a> |
| **PowerPoint** <sup>365/2021/2024</sup> **Standalone Installer** | `{get_package_detail(packages, 'PowerPoint', 'short_version')}` | com.microsoft.powerpoint | <a href="https://go.microsoft.com/fwlink/?linkid=525136"><img src="images/PPT3_512x512x32.png" alt="Download Image" width="80"></a> |
| **Outlook** <sup>365/2021/2024</sup> **Standalone Installer** | `{get_package_detail(packages, 'Outlook', 'short_version')}` | com.microsoft.outlook | <a href="https://go.microsoft.com/fwlink/?linkid=2228621"><img src="images/Outlook_512x512x32.png" alt="Download Image" width="80"></a>|
| **OneNote** <sup>365/2021/2024</sup> **Standalone Installer** | `{get_package_detail(packages, 'OneNote', 'short_version')}` | com.microsoft.onenote.mac | <a href="https://go.microsoft.com/fwlink/?linkid=820886"><img src="images/OneNote_512x512x32.png" alt="Download Image" width="80"></a> |
| **OneDrive Standalone Installer**<br><a href="https://support.microsoft.com/en-us/office/onedrive-release-notes-845dcf18-f921-435e-bf28-4e24b95e5fc0#OSVersion=Mac" style="text-decoration: none;"><small>_Release Notes_</small></a> | `{get_package_detail(packages, 'OneDrive', 'short_version')}` | com.microsoft.OneDrive | <a href="https://go.microsoft.com/fwlink/?linkid=823060"><img src="images/OneDrive_512x512x32.png" alt="Download Image" width="80"></a> |
| **Skype for Business Standalone Installer**<br><a href="https://support.microsoft.com/en-us/office/follow-the-latest-updates-in-skype-for-business-cece9f93-add1-4d93-9a38-56cc598e5781?ui=en-us&rs=en-us&ad=us" style="text-decoration: none;"><small>_Release Notes_</small></a>  | `{get_package_detail(packages, 'Skype', 'short_version')}` | com.microsoft.SkypeForBusiness | <a href="{get_package_detail(packages, 'Skype', 'update_download')}"><img src="images/skype_for_business.png" alt="Download Image" width="80"></a> |
| **Teams Standalone Installer**<br><a href="https://support.microsoft.com/en-us/office/what-s-new-in-microsoft-teams-d7092a6d-c896-424c-b362-a472d5f105de" style="text-decoration: none;"><small>_Release Notes_</small></a>  | `{get_package_detail(packages, 'Teams', 'short_version')}` | com.microsoft.teams2 | <a href="https://go.microsoft.com/fwlink/?linkid=2249065"><img src="images/teams_512x512x32.png" alt="Download Image" width="80"></a> |
| **InTune Company Portal Standalone Installer**<br><a href="https://aka.ms/intuneupdates" style="text-decoration: none;"><small>_Release Notes_</small></a> | `{get_package_detail(packages, 'Intune', 'short_version')}` | com.microsoft.CompanyPortalMac | <a href="https://go.microsoft.com/fwlink/?linkid=853070"><img src="images/companyportal.png" alt="Download Image" width="80"></a> |
| **Edge Standalone Installer** <sup>_(Stable Channel)_</sup><br><a href="https://learn.microsoft.com/en-us/deployedge/microsoft-edge-relnote-stable-channel" style="text-decoration: none;"><small>_Release Notes_</small></a>| `{get_package_detail(packages, 'Edge', 'short_version')}` | com.microsoft.edgemac | <a href="https://go.microsoft.com/fwlink/?linkid=2093504"><img src="images/edge_app.png" alt="Download Image" width="80"></a>|
| **Defender for Endpoint Installer**<br><a href="https://learn.microsoft.com/microsoft-365/security/defender-endpoint/mac-whatsnew" style="text-decoration: none;"><small>_Release Notes_</small></a> | `{get_package_detail(packages, 'Defender For Endpoint', 'short_version')}` | com.microsoft.wdav | <a href="https://go.microsoft.com/fwlink/?linkid=2097502"><img src="images/defender_512x512x32.png" alt="Download Image" width="80"></a> |
| **Defender for Consumers Installer**<br><a href="https://learn.microsoft.com/microsoft-365/security/defender-endpoint/mac-whatsnew" style="text-decoration: none;"><small>_Release Notes_</small></a> | `{get_package_detail(packages, 'Defender For Consumers', 'short_version')}` | com.microsoft.wdav | <a href="https://go.microsoft.com/fwlink/?linkid=2247001"><img src="images/defender_512x512x32.png" alt="Download Image" width="80"></a> |
| **Defender SHIM Installer** | `{get_package_detail(packages, 'Defender Shim', 'short_version')}` | com.microsoft.wdav.shim | <a href="{get_package_detail(packages, 'Defender Shim', 'update_download')}"><img src="images/defender_512x512x32.png" alt="Download Image" width="80"></a> |
| **Windows App Standalone Installer** </a><sup>_(Remote Desktop <img src="images/microsoft-remote-desktop-logo.png" alt="Remote Desktop" width="15"></a>)_</sup><br><a href="https://learn.microsoft.com/en-us/windows-app/whats-new?tabs=macos" style="text-decoration: none;"><small>_Release Notes_</small> | `{get_package_detail(packages, 'Windows App', 'short_version')}` | com.microsoft.rdc.macos | <a href="https://go.microsoft.com/fwlink/?linkid=868963"><img src="images/windowsapp.png" alt="Download Image" width="80"></a> |
| **Visual Studio Code Standalone Installer**<br><a href="https://code.visualstudio.com/updates/" style="text-decoration: none;"><small>_Release Notes_</small></a>  | `{get_package_detail(packages, 'Visual', 'short_version')}` | com.microsoft.VSCode | <a href="https://go.microsoft.com/fwlink/?linkid=2156837"><img src="images/Code_512x512x32.png" alt="Download Image" width="80"></a>|
| **AutoUpdate Standalone Installer**<br><a href="https://learn.microsoft.com/en-us/officeupdates/release-history-microsoft-autoupdate" style="text-decoration: none;"><small>_Release Notes_</small></a>  | `{get_package_detail(packages, 'MAU', 'short_version')}` | com.microsoft.autoupdate | <a href="https://go.microsoft.com/fwlink/?linkid=830196"><img src="images/autoupdate.png" alt="Download Image" width="80"></a>|
| **Licensing Helper Tool Installer** | `{get_package_detail(packages, 'Licensing Helper Tool', 'short_version')}` | N/A | <a href="{get_package_detail(packages, 'Licensing Helper Tool', 'latest_download')}"><img src="images/pkg-icon.png" alt="Download Image" width="80"></a>|
| **Quick Assist Installer** | `{get_package_detail(packages, 'Quick Assist', 'short_version')}` | com.microsoft.quickassist | <a href="{get_package_detail(packages, 'Quick Assist', 'latest_download')}"><img src="images/quickassist.png" alt="Download Image" width="80"></a>|
| **Remote Help Installer** | `{get_package_detail(packages, 'Remote Help', 'short_version')}` | com.microsoft.remotehelp | <a href="{get_package_detail(packages, 'Remote Help', 'latest_download')}"><img src="images/remotehelp.png" alt="Download Image" width="80"></a>|

<sup>_**For items without specific release notes, please refer to the release notes for the entire suite.**_</sup> <br>

<sup>_**All apps include MAU with installation, except for Skype for Business, OneDrive, Defender SHIM, Licensing Helper Tool, Quick Assist, and Remote Help.**_</sup>

| **Product Package** | **Link** | **<img src="images/sha-256.png" alt="Download Image" width="20">SHA256 Hash<img src="images/sha-256.png" alt="Download Image" width="20">** |
|----------------------|----------|------------------|
| **Microsoft** <sup>365/2021/2024</sup> **and Office Suite Installer**<br><sub>_(Includes Word, Excel, PowerPoint, Outlook, OneNote, OneDrive, and MAU)_</sub> | <a href="https://go.microsoft.com/fwlink/?linkid=525133"><img src="images/suite.png" alt="Download Image" width="60"></a> | `{get_package_detail(packages, 'Microsoft Office Suite', 'sha256')}` |
| **Microsoft** <sup>365/2021/2024</sup> **BusinessPro Suite Installer**<br><sub>_(Includes Word, Excel, PowerPoint, Outlook, OneNote, OneDrive, Teams, Defender Shim, and MAU)_</sub> | <a href="https://go.microsoft.com/fwlink/?linkid=2009112"><img src="images/suite.png" alt="Download Image" width="60"></a> | `{get_package_detail(packages, 'Microsoft BusinessPro Suite', 'sha256')}` |
| **Word** <sup>365/2021/2024</sup> **Standalone Installer** | <a href="https://go.microsoft.com/fwlink/?linkid=525134"><img src="images/MSWD_512x512x32.png" alt="Download Image" width="60"></a> | `{get_package_detail(packages, 'Word', 'sha256')}` |
| **Excel** <sup>365/2021/2024</sup> **Standalone Installer** | <a href="https://go.microsoft.com/fwlink/?linkid=525135"><img src="images/XCEL_512x512x32.png" alt="Download Image" width="60"></a> | `{get_package_detail(packages, 'Excel', 'sha256')}` |
| **PowerPoint** <sup>365/2021/2024</sup> **Standalone Installer** | <a href="https://go.microsoft.com/fwlink/?linkid=525136"><img src="images/PPT3_512x512x32.png" alt="Download Image" width="60"></a> | `{get_package_detail(packages, 'PowerPoint', 'sha256')}` |
| **Outlook** <sup>365/2021/2024</sup> **Standalone Installer**| <a href="https://go.microsoft.com/fwlink/?linkid=525137"><img src="images/Outlook_512x512x32.png" alt="Download Image" width="60"></a> | `{get_package_detail(packages, 'Outlook', 'sha256')}` |
| **OneNote** <sup>365/2021/2024</sup> **Standalone Installer** | <a href="https://go.microsoft.com/fwlink/?linkid=820886"><img src="images/OneNote_512x512x32.png" alt="Download Image" width="60"></a> | `{get_package_detail(packages, 'OneNote', 'sha256')}` |
| **OneDrive Standalone Installer** | <a href="https://go.microsoft.com/fwlink/?linkid=823060"><img src="images/OneDrive_512x512x32.png" alt="Download Image" width="60"></a> | `{get_package_detail(packages, 'OneDrive', 'sha256')}` |
| **Skype for Business Standalone Installer** | <a href="{get_package_detail(packages, 'Skype', 'update_download')}"><img src="images/skype_for_business.png" alt="Download Image" width="60"></a> | `{get_package_detail(packages, 'Skype', 'sha256')}` |
| **Teams Standalone Installer** | <a href="https://go.microsoft.com/fwlink/?linkid=2249065"><img src="images/teams_512x512x32.png" alt="Download Image" width="60"></a> | `{get_package_detail(packages, 'Teams', 'sha256')}` |
| **InTune Company Portal Standalone Installer** | <a href="https://go.microsoft.com/fwlink/?linkid=853070"><img src="images/companyportal.png" alt="Download Image" width="60"></a> | `{get_package_detail(packages, 'Intune', 'sha256')}` |
| **Edge Standalone Installer** <sup>_(Stable Channel)_</sup> | <a href="https://go.microsoft.com/fwlink/?linkid=2093504"><img src="images/edge_app.png" alt="Download Image" width="60"></a> | `{get_package_detail(packages, 'Edge', 'sha256')}` |
| **Defender For Endpoint Installer** | <a href="https://go.microsoft.com/fwlink/?linkid=2097502"><img src="images/defender_512x512x32.png" alt="Download Image" width="60"></a> | `{get_package_detail(packages, 'Defender For Endpoint', 'sha256')}` |
| **Defender For Consumer Installer** | <a href="https://go.microsoft.com/fwlink/?linkid=2097001"><img src="images/defender_512x512x32.png" alt="Download Image" width="60"></a> | `{get_package_detail(packages, 'Defender For Consumers', 'sha256')}` |
| **Defender Shim Installer** | <a href="{get_package_detail(packages, 'Defender Shim', 'latest_download')}"><img src="images/defender_512x512x32.png" alt="Download Image" width="60"></a> | `{get_package_detail(packages, 'Defender Shim', 'sha256')}` |
| **Windows App Standalone Installer** | <a href="https://go.microsoft.com/fwlink/?linkid=868963"><img src="images/windowsapp.png" alt="Download Image" width="60"></a> | `{get_package_detail(packages, 'Windows App', 'sha256')}` |
| **Visual Studio Code Standalone Installer** | <a href="https://go.microsoft.com/fwlink/?linkid=2156837"><img src="images/Code_512x512x32.png" alt="Download Image" width="60"></a> | `{get_package_detail(packages, 'Visual', 'sha256')}` |
| **AutoUpdate Standalone Installer** | <a href="https://go.microsoft.com/fwlink/?linkid=830196"><img src="images/autoupdate.png" alt="Download Image" width="60"></a> | `{get_package_detail(packages, 'MAU', 'sha256')}` |
| **Licensing Helper Tool Installer** | <a href="{get_package_detail(packages, 'Licensing Helper Tool', 'latest_download')}"><img src="images/pkg-icon.png" alt="Download Image" width="60"></a> | `{get_package_detail(packages, 'Licensing Helper Tool', 'sha256')}` |
| **Quick Assist Installer** | <a href="{get_package_detail(packages, 'Quick Assist', 'latest_download')}"><img src="images/quickassist.png" alt="Download Image" width="60"></a> | `{get_package_detail(packages, 'Quick Assist', 'sha256')}` |
| **Remote Help Installer** | <a href="{get_package_detail(packages, 'Remote Help', 'latest_download')}"><img src="images/remotehelp.png" alt="Download Image" width="60"></a> | `{get_package_detail(packages, 'Remote Help', 'sha256')}` |

<sup>_<img src="images/sha-256.png" alt="Download Image" width="15">[**How to Get the SHA256 Guide**](/guides/How_To_SHA256.md)<img src="images/sha-256.png" alt="Download Image" width="15">_</sup>

| **Special Product Package** | **CFBundle Version** | **MAU Status** | **Download** |
|----------------------|----------------------|--------------------------|--------------|
| **Word** <sup>365/2021/2024</sup> **</sup> Standalone Installer** | `{get_package_detail(packages, 'Word', 'short_version')}` | No MAU | <a href="{get_package_detail(packages, 'Word', 'update_download')}"><img src="images/MSWD_512x512x32.png" alt="Download Image" width="80"></a> |
| **Excel** <sup>365/2021/2024</sup> **Standalone Installer** | `{get_package_detail(packages, 'Excel', 'short_version')}` | No MAU | <a href="{get_package_detail(packages, 'Excel', 'update_download')}"><img src="images/XCEL_512x512x32.png" alt="Download Image" width="80"></a> |
| **PowerPoint** <sup>365/2021/2024</sup> **Standalone Installer** | `{get_package_detail(packages, 'PowerPoint', 'short_version')}` | No MAU | <a href="{get_package_detail(packages, 'PowerPoint', 'update_download')}"><img src="images/PPT3_512x512x32.png" alt="Download Image" width="80"></a> |
| **Outlook** <sup>365/2021/2024</sup> **Standalone Installer**<sup>_(Weekly Channel)_</sup>| `{get_package_detail(packages, 'Outlook', 'short_version')}` | No MAU | <a href="{get_package_detail(packages, 'Outlook', 'update_download')}"><img src="images/Outlook_512x512x32.png" alt="Download Image" width="80"></a>|
| **OneNote** <sup>365/2021/2024</sup> **Standalone Installer** | `{get_package_detail(packages, 'OneNote', 'short_version')}` | No MAU | <a href="{get_package_detail(packages, 'OneNote', 'update_download')}"><img src="images/OneNote_512x512x32.png" alt="Download Image" width="80"></a> |
| **InTune Company Portal Standalone Installer**<br><a href="https://aka.ms/intuneupdates" style="text-decoration: none;"><small>_Release Notes_</small></a> | `{get_package_detail(packages, 'Intune', 'short_version')}` | No MAU | <a href="{get_package_detail(packages, 'Intune', 'update_download')}"><img src="images/companyportal.png" alt="Download Image" width="80"></a> |
| **Outlook** <sup>365/2021/2024</sup> **Standalone Installer**<sup>_(Monthly Channel)_</sup>| `N/A - Check Release Notes` | Contains MAU | <a href="https://go.microsoft.com/fwlink/?linkid=525137"><img src="images/Outlook_512x512x32.png" alt="Download Image" width="80"></a>|


|      Update History                   |          Microsoft Update Channels               |
|-------------------------|-------------------------|
| <img src="images/Microsoft_Logo_512px.png" alt="Download Image" width="20"> [Microsoft 365/2021/2024](https://learn.microsoft.com/en-us/officeupdates/update-history-office-for-mac) | <img src="images/Microsoft_Logo_512px.png" alt="Download Image" width="20">  [Microsoft 365 Apps](https://learn.microsoft.com/en-us/microsoft-365-apps/updates/overview-update-channels) |


## **<img src="images/repair.png" alt="Repair Image" width="20"></a> Microsoft Office Repair Tools <img src="images/repair.png" alt="Repair Image" width="20"></a>**

### **<img src="/images/Office_Reset_512x512.png" alt="Office Reset Logo" width="25"> [Office-Reset.com](https://office-reset.com/macadmins/)<img src="/images/Office_Reset_512x512.png" alt="Office Reset Logo" width="25">**
- A free tool to fix issues with Microsoft Office apps on macOS (e.g., crashes, performance problems). It offers various packages for resetting settings and clearing cache.
<br>**_<img src="images/warning.png" alt="Warning Logo" width="25"> No Longer Updated/Maintained <img src="images/warning.png" alt="Warning Logo" width="25">_**

### **<img src="images/pkg-icon.png" alt="Pkg Logo" width="25"> [Office-Reset Packages](/office_reset_pkgs/)<img src="images/pkg-icon.png" alt="Pkg Logo" width="25">**
- Archived copies of the original Office-Reset packages, now saved to this repository.
<br>**_<img src="images/warning.png" alt="Warning Logo" width="25"> No Longer Updated/Maintained <img src="images/warning.png" alt="Warning Logo" width="25">_**

### **<img src="images/repair.png" alt="Repair Image" width="20"></a> [Office Reset Scripts](/office_reset_scripts/) <img src="images/repair.png" alt="Repair Image" width="20"></a>**
- Fork and update these scripts from the original packages on office-reset.com.
<br>**_<img src="images/community.png" alt="Community Logo" width="25"> Needs More Community Contributions <img src="images/community.png" alt="Community Logo" width="25">_**

## **<img src="images/script.png" alt="Script Image" width="20"> Microsoft Scripts <img src="images/script.png" alt="Script Image" width="20">**

These scripts automate the process of downloading, installing, updating, and managing Microsoft products.

- **Download & Install Microsoft Products**: [View Script](https://gist.github.com/talkingmoose/b6637160b65b751824943ede022daa17) by [TalkingMoose](https://github.com/talkingmoose)  
  This script automates the downloading and installation of the latest Microsoft products using direct links and includes optional SHA256 verification for added security.

- **Install Office 365 Pro**: [View Script](https://github.com/microsoft/shell-intune-samples/blob/master/macOS/Apps/Office%20for%20Mac/installOffice365Pro.sh) by [Microsoft](https://github.com/microsoft)  
  This script automates the downloading and installation of the Office 365 Pro.

- **Installomator**: [View Script](https://github.com/Installomator/Installomator) by [Installomator](https://github.com/Installomator)  
  A powerful tool for automating the deployment of Microsoft Office products on macOS, simplifying downloading, installation, and updates.

- **Various Tools**: [View Scripts](https://github.com/pbowden-msft?tab=repositories) by [pbowden-msft](https://github.com/pbowden-msft)  
  A collection of tools for repairing, setting up, and automating Microsoft Office products.

## **Microsoft Office Preference Keys**

PLIST (Property List) files are used by macOS to store settings and preferences for apps, services, and system configurations, allowing Mac admins to:

- **Customize deployments**
- **Enforce policies**
- **Manage application behavior efficiently**

For a detailed guide on how to create and manage PLIST files, refer to the [How to Plist Guide](/guides/How_To_plist.md).

### **Recommended Resources:**

#### **<img src="images/MAF_Badge_4c.png" alt="Download Image" width="30"> Mac Admin Community-Driven Preferences List (Highly Recommended!)**:
- [View Google Doc](https://docs.google.com/spreadsheets/d/1ESX5td0y0OP3jdzZ-C2SItm-TUi-iA_bcHCBvaoCumw/edit?gid=0#gid=0)

#### **<img src="images/Microsoft_Logo_512px.png" alt="Download Image" width="20"> Official Microsoft Documentation:**

- [General PLIST Preferences](https://learn.microsoft.com/en-us/microsoft-365-apps/mac/deploy-preferences-for-office-for-mac)  
- [App-Specific Preferences](https://learn.microsoft.com/en-us/microsoft-365-apps/mac/set-preference-per-app)  
- [Outlook Preferences](https://learn.microsoft.com/en-us/microsoft-365-apps/mac/preferences-outlook)  
- [Office Suite Preferences](https://learn.microsoft.com/en-us/microsoft-365-apps/mac/preferences-office)

## **Contributing and Providing Feedback**

We warmly welcome your contributions and feedback to **macadmins_msft**! Here’s how you can get involved:

### 📋 **Report Issues**  
Have a bug to report or a feature to request? Submit an issue on our [GitHub Issues page](https://github.com/cocopuff2u/macadmins_msft/issues).  

### 💬 **Join the Discussion**  
Connect and collaborate in the [GitHub Discussions](https://github.com/cocopuff2u/macadmins_msft/discussions) or the [Mac Admins Slack Channel](https://macadmins.slack.com/).  
- **Reach Out Directly:** Contact me on Slack at `cocopuff2u` for direct collaboration or questions.
- **New to Slack?** [Sign up here](https://join.slack.com/t/macadmins/shared_invite/zt-2tq3md5zr-jDtuUFHAFa8CIBwPhpFfFQ).  
- **Existing User?** [Sign in here](https://macadmins.slack.com/).  
- **Explore Slack Channels:**  
    - `#microsoft-office`  
    - `#microsoft-autoupdate`  
    - `#microsoft-intune`  
    - `#microsoft-windows-app`  
    - `#microsoft-office-365`  
    - `#microsoft-teams`  

### ✉️ **Contact via Email**  
For inquiries, reach out directly at [cocopuff2u@yahoo.com](mailto:cocopuff2u@yahoo.com).  

### 🛠️ **Contribute Directly**  
Fork the repository, make your changes, and submit a pull request—every contribution counts!  

### 💡 **Share Your Feedback**  
Help us improve! Share your ideas and suggestions in the [GitHub Discussions](https://github.com/cocopuff2u/macadmins_msft/discussions) or via email.

### 🌟 **Support the Project**  
Any contributions to this project go toward building a fully-featured URL for this site or are shared with the Mac Admins community to keep information accessible for everyone. If you’re feeling generous, you can also support the project by donating to my coffee fund. 

## **Helpful Links**

Below are a list of helpful links.
- **Microsoft Versioning Shenanigans**: [View Link](https://macmule.com/2018/09/24/microsoft-office-for-mac-changes-versioning-shenanigans/)
- **Microsoft Deployment Options**: [View Link](https://learn.microsoft.com/en-us/microsoft-365-apps/mac/deployment-options-for-office-for-mac)
- **Microsoft Deploy From App Store**: [View Link](https://learn.microsoft.com/en-us/microsoft-365-apps/mac/deploy-mac-app-store)
- **JAMF Technical Paper: Managing Microsoft Office**: [View Link](https://learn.jamf.com/en-US/bundle/technical-paper-microsoft-office-current/page/User_Experience_Configuration.html)

## **Trademarks**

- **Microsoft 365, Office 365, Excel, PowerPoint, Outlook, OneDrive, OneNote, Teams** are trademarks of Microsoft Corporation.
- **Mac** and **macOS** are trademarks of Apple Inc.
- Other names and brands may be claimed as the property of their respective owners.
"""
    logging.info("README content generated successfully")
    
    return content

def overwrite_readme(file_path, content):
    with open(file_path, "w") as file:
        file.write(content)
    print(f"README.md has been overwritten.")

def get_package_detail(packages, package_name, detail):
    # Ensure package name and detail are lowercase for easy matching
    package_name = package_name.lower()
    detail = detail.lower()

    # Check if package and detail exist
    if package_name in packages and detail in packages[package_name]:
        return packages[package_name][detail]
    else:
        return None

if __name__ == "__main__":
    # Define file paths
    xml_file_path = "latest.xml"  # Update this path if the file is located elsewhere
    readme_file_path = "README.md"

    # Parse the XML and generate content
    global_last_updated, packages = parse_latest_xml(xml_file_path)
    readme_content = generate_readme_content(global_last_updated, packages)

    # Overwrite the README file
    overwrite_readme(readme_file_path, readme_content)
