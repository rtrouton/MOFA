# **Guide: Creating and Deploying .plist Files with MDM (Jamf Pro or Intune)**
Learn how to create `.plist` files, store them in the correct locations, and deploy them using MDM solutions like **Jamf Pro** or **Intune**. This guide will use Outlook as the example.

### **1. Understanding and Creating `.plist` Files**

In macOS, `.plist` (Property List) files are used to store application configuration settings. These files allow administrators to define preferences, such as enabling or disabling features, to customize user experiences in Microsoft Outlook. These configurations can be deployed efficiently using a Mobile Device Management (MDM) solution, such as **Jamf Pro** or **Intune**, by converting `.plist` files into configuration profiles.

#### **Creating a .plist File from Scratch:**

To create a `.plist` file from scratch for configuring Microsoft Outlook preferences, follow these steps:

1. **Open a Text Editor:** Use any text editor, such as `TextEdit` in plain text mode, or a code editor like `VS Code`.
2. **Add XML Structure:**
   - Start with the XML declaration and a root `<plist>` element.
   - Inside the `<dict>` tag, define the key-value pairs that represent the configuration settings.

   **Example:**
   ```xml
   <?xml version="1.0" encoding="UTF-8"?>
   <!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
   <plist version="1.0">
       <dict>
           <key>PlayMyEmailsEnabled</key>
           <false/>
       </dict>
   </plist>
   ```

   - **Key:** `PlayMyEmailsEnabled` controls the "Play My Emails" feature in Outlook.
   - **Value:** `false` disables the feature.

3. **Save the File:**
   Name the file with a `.plist` extension, such as `com.microsoft.Outlook.plist`. Each application or suite will have its own required name.

---

### **2. Location of .plist Files on macOS**

The location of `.plist` files depends on whether the preferences are user-specific or managed by MDM.

- **For User Preferences (Local, Unmanaged):**
  ```plaintext
  ~/Library/Preferences/com.microsoft.Outlook.plist
  ```
  This file contains the user-specific settings for Outlook.

- **For Managed Preferences (Deployed via MDM):**
  ```plaintext
  /Library/Managed Preferences/com.microsoft.Outlook.plist
  ```
  MDM configurations are applied here, overriding user-specific preferences.

---

### **3. Deploying via MDM: Jamf Pro or Intune**

Once you've created your `.plist` file, you can deploy it using **Jamf Pro** or **Intune**.

#### **Option 1: Jamf Pro**
1. **Create a Configuration Profile:**
   - Log in to **Jamf Pro**.
   - Navigate to **Configuration Profiles** > **New**.
   - Select the **Custom Settings** payload.
   - Upload the `.plist` file or paste its XML content into the **Custom Settings** field.

2. **Assign Scope:**
   - Assign the profile to devices or users using Smart Groups or Static Groups.

3. **Deploy:**
   - Save the profile. Devices will apply the configuration during the next MDM sync.

---

#### **Option 2: Microsoft Intune**
1. **Prepare the Configuration Profile:**
   - Wrap the `.plist` file into a `.mobileconfig` profile. Tools like **Profile Creator** can help with this.

2. **Upload to Intune:**
   - Log in to **Microsoft Endpoint Manager Admin Center**.
   - Go to **Devices** > **macOS** > **Configuration Profiles** > **Create Profile**.
   - Choose **Templates** > **Custom** and upload the `.mobileconfig` file.

3. **Assign Scope:**
   - Assign the profile to the appropriate user or device groups.

4. **Deploy:**
   - Save and assign the profile. Devices will sync and apply the configuration during their next check-in.

---

### **4. Helpful Resources for Configuration Options**

For detailed guidance on available preferences and settings, refer to these valuable resources. We recommend starting with the Mac Admin Community's comprehensive, crowd-sourced list:

- **Mac Admin Community-Driven Preferences List:**
  - [View Google Doc](https://docs.google.com/spreadsheets/d/1ESX5td0y0OP3jdzZ-C2SItm-TUi-iA_bcHCBvaoCumw/edit?gid=0#gid=0)

- **Microsoft Documentation for macOS Preferences:**
  - [General Plist Preferences](https://learn.microsoft.com/en-us/microsoft-365-apps/mac/deploy-preferences-for-office-for-mac)
  - [App-Specific Preferences](https://learn.microsoft.com/en-us/microsoft-365-apps/mac/set-preference-per-app)
  - [Outlook Preferences](https://learn.microsoft.com/en-us/microsoft-365-apps/mac/preferences-outlook)
  - [Office Suite Preferences](https://learn.microsoft.com/en-us/microsoft-365-apps/mac/preferences-office)

These resources provide detailed `.plist` keys, values, and usage examples to help tailor your deployment.

---

### **5. Test and Validate**

After deploying the `.plist` file, validate the configuration to ensure the settings were applied successfully:

1. **Check Managed Preferences:**
   On the target Mac, check for the applied configuration in:
   ```bash
   /Library/Managed Preferences/com.microsoft.Outlook.plist
   ```

2. **Inspect Settings with `defaults`:**
   Use the `defaults` command to verify the applied preference:
   ```bash
   defaults read com.microsoft.Outlook PlayMyEmailsEnabled
   ```

3. **Review Logs:**
   - **Jamf Pro:** Check the deployment status in the Jamf Pro dashboard.
   - **Intune:** Review deployment reports in the Endpoint Manager Admin Center.

---

### **6. Tips for Success**
- **Microsoft Documentation:** Always refer to Microsoft's official preference keys for Outlook on macOS.
  Example: [Office for Mac Preferences](https://learn.microsoft.com/en-us/microsoft-365-apps/mac/deploy-preferences-for-office-for-mac)
- **Test in Small Groups:** Before deploying organization-wide, test profiles on a small group of devices to ensure compatibility.
- **Bundle Identifier:** Make sure the profile targets the correct app. For Outlook, the identifier is `com.microsoft.Outlook`.

---

### **7. Troubleshooting Common Issues**

Adding a troubleshooting section will help users address any common issues they might encounter when deploying `.plist` files with MDM.

- **Plist Not Applying:**
   - Ensure that the `.plist` file is located in the correct directory (e.g., `/Library/Managed Preferences/` for MDM deployments).
   - Verify that the configuration profile is assigned to the correct target devices or users.

- **Outlook Settings Not Reflecting:**
   - Use the `defaults` command to check if the setting has been applied. If not, ensure that the correct keys are being used in the `.plist`.
   - Check for conflicts with user-specific preferences located in `~/Library/Preferences/`.

- **MDM Sync Issues:**
   - Ensure the device has checked in with the MDM server. You can force a sync from the **Jamf Pro** or **Intune** console or manually on the Mac.
   - Check MDM logs for any errors during profile deployment.

---

### **Additional Resources and Links**

1. **Apple Developer Documentation:**
   - [Property List Programming Guide](https://developer.apple.com/library/archive/documentation/Cocoa/Conceptual/PropertyLists/Introduction/Introduction.html)
     This provides an in-depth understanding of `.plist` structure, serialization, and usage.

2. **Jamf Pro Knowledge Base:**
   - [Jamf Pro Documentation](https://jamf.com/resources/)
     A great resource for learning more about Jamf Pro’s configuration profiles and the process for managing macOS devices.

3. **Intune for Education Resources:**
   - [Microsoft Intune Documentation](https://learn.microsoft.com/en-us/mem/intune/)
     A comprehensive guide to managing Apple devices in Microsoft Intune.

4. **Profile Creator Tool for macOS:**
   - [Profile Creator](https://github.com/ProfileCreator/ProfileCreator)
     An open-source tool for creating `.mobileconfig` profiles from `.plist` files.

5. **Plist Editor Tools:**
   - [PlistEdit Pro](https://www.fatcatsoftware.com/plisteditpro/)
     A more advanced graphical tool for editing `.plist` files on macOS.
