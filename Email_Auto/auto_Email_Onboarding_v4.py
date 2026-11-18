
from hashlib import new
import requests
import json
import os
import pandas as pd
import win32com.client as win32
from dotenv import load_dotenv
from pathlib import Path
load_dotenv(Path(__file__).resolve().parent / ".env")

client_id = os.getenv("CLIENT_ID")
client_secret = os.getenv("CLIENT_SECRET")

if not client_id or not client_secret:
    raise ValueError("CLIENT_ID or CLIENT_SECRET is missing from .env")
else:
    print("CLIENT_ID and CLIENT_SECRET loaded successfully.")

base_url = "https://prod.api.nvidia.com/sys/snow/api/now/v1/table"

Emea_Engineers =[
    "4f0620a897db47144962b5c3f153af86",
    "7efe6fc9db34d1505aa1ef92ca961902",
    "b23f6b4ddb34d1505aa1ef92ca9619bf",
    "28823745dbb4d1505aa1ef92ca96195b",
    "cbffabc1db74d1505aa1ef92ca96190c",
    "52b45424476b0f94f5c498fb736d43bb",
    "867fa701db74d1505aa1ef92ca9619c9",
    "09823745dbb4d1505aa1ef92ca9619b4",
    "b8003fc1db74d1505aa1ef92ca96191a",
    "d86377c9dbb4d1505aa1ef92ca96194c",
    "5e037b09dbb4d1505aa1ef92ca961902"

 ]



headers = {
    "Accept": "application/json"
}

params = {
    "sysparm_limit": 100,
    "sysparm_offset": 0,
    "sysparm_query": (
        f"assigned_toIN{','.join(Emea_Engineers)}"
        "^sys_class_name=sc_task"
        "^assignment_group.name=NV IT Onboarding"
        "^item.name=Onboarding General-Purpose Computer"
        "^short_description=setup/deployment Item"
        "^state=3"
        "^closed_atRELATIVEGT@dayofweek@ago@10"
        #"^cmdb_ciISNOTEMPTY"
    ),
    "sysparm_fields": (
        "sys_id,number,short_description,assigned_to,"
        "assignment_group,u_task_requested_for,requested_for,"
        "cmdb_ci,asset,location,request_item"
    ),
    "sysparm_display_value": "all"
}

# Retrieve all matching tasks
all_tasks = []
tasks_url = f"{base_url}/task"

while True:
    response = requests.get(
        tasks_url,
        params=params,
        headers=headers,
        auth=(client_id, client_secret),
        timeout=(10, 120)
    )

    # print("Status Code:", response.status_code)
    # print("Response:", response.text)
    # response.raise_for_status()

    task_details = response.json().get("result", {})

    response.raise_for_status()
    tasks = response.json().get("result", [])

    if not tasks:
        break

    all_tasks.extend(tasks)

    #print(f"Retrieved {len(tasks)} tasks (total: {len(all_tasks)})")

    if len(tasks) < params["sysparm_limit"]:
        break

    params["sysparm_offset"] += params["sysparm_limit"]
#print (all_tasks)
print(f"\nTotal SCTASKs: {len(all_tasks)}")
#print (all_tasks)



# Store details for every task
onboarding_details = []

for task in all_tasks:

    # Task details
    task_sys_id = task.get("sys_id", {}).get("display_value", "")
    task_number = task.get("number", {}).get("display_value", "")
    asset = ""
    ritm = task.get("request_item", {})

    if isinstance(ritm, dict):
     ritm_sys_id = ritm.get("value", "")
    else:
     ritm_sys_id = ritm

    # Get the requested employee reference
    response = requests.get(
        f"{base_url}/sc_task/{task_sys_id}",
        params={
            "sysparm_fields": "number,u_task_requested_for,cmdb_ci,asset",
            "sysparm_display_value": "all"
        },
        headers=headers,
        auth=(client_id, client_secret),
        timeout=(10, 120)
    )

    response.raise_for_status()
    task_details = response.json().get("result", {})

    user_sys_id = (
        task_details.get("u_task_requested_for", {}).get("value", "")
    )
 
    if not user_sys_id:
        print(f"Skipping {task_number}: no requested employee found")
        continue
    #print(f"{task_number}: {user_sys_id}")

    # Get the employee's assigned assets and exclude accessories
    asset = ""
    asset_tag = ""
    asset_sys_id = ""

    response = requests.get(
        f"{base_url}/alm_asset",
        params={
            "sysparm_query": f"assigned_to={user_sys_id}",
            "sysparm_fields": "sys_id,model,display_name,asset_tag,serial_number",
            "sysparm_display_value": "all",
            "sysparm_limit": 100
        },
        headers=headers,
        auth=(client_id, client_secret),
        timeout=(10, 120)
    )

    response.raise_for_status()
    assets = response.json().get("result", [])

    for item in assets:
        model = item.get("model", {})
        model_name = model.get("display_value", "") if isinstance(model, dict) else model

        if (
          "backpack" not in model_name.lower()
          and "logitech" not in model_name.lower()
          and "p3424" not in model_name.lower()
          and "p2724" not in model_name.lower()
          and "p3426" not in model_name.lower()
          and "yubikey" not in model_name.lower()
          and "startech" not in model_name.lower()
           ):
            display = item.get("display_name", {})
            asset = display.get("display_value", "") if isinstance(display, dict) else display

            sys_id = item.get("sys_id", {})
            asset_sys_id = sys_id.get("value", "") if isinstance(sys_id, dict) else sys_id

            tag = item.get("asset_tag", {})
            asset_tag = tag.get("display_value", "") if isinstance(tag, dict) else tag
            break


 # Get employee details
    response = requests.get(
        f"{base_url}/sys_user/{user_sys_id}",
        params={
            "sysparm_fields": "sys_id,name,email,location,user_name",
            "sysparm_display_value": "all"
        },
        headers=headers,
        auth=(client_id, client_secret),
        timeout=(10, 120)
    )

    response.raise_for_status()
    user = response.json().get("result", {})

    employee_name = user.get("name", {}).get("display_value", "")
    email = user.get("email", {}).get("value", "")
    employee_location = user.get("location", {}).get("display_value", "")
    userid = user.get("user_name", {}).get("value", "")
    #print(userid)

    # user_sys_id = (
    #     task_details.get("u_task_requested_for", {}).get("value", "")
    # )

    # Get onboarding date records for this employee
    response = requests.get(
        #f"{base_url}/u_onboarding_changes",
        f"{base_url}/sys_user",
        params={
            "sysparm_query": (
                #f"u_user_profile={user_sys_id}"
                f"sys_id={user_sys_id}"
                             ),
            "sysparm_fields": (
                "sys_id,u_user_profile,u_onboarding_hire_date,sys_updated_on,user_name"
            ),
            "sysparm_display_value": "all",
            "sysparm_limit": 100
        },
        headers=headers,
        auth=(client_id, client_secret),
        timeout=(10, 120)
    )
  ##############

    #print("HTTP status:", response.status_code)
    #print("Response body:", response.text)
    response.raise_for_status()
    records = response.json().get("result", [])
    #print(f"Records: {records}")

    onboarding_date = ""

    if records:
        onboarding_date = records[0].get(
            "u_onboarding_hire_date", {}        ).get("display_value", "")

    # Save all details together for later calculations and email creation
    onboarding_record = {
        "task_number": task_number,
        "task_sys_id": task_sys_id,
        "employee_name": employee_name,
        "email": email,
        "location": employee_location,
        "onboarding_date": onboarding_date,
        "asset": asset,
        "asset_tag": asset_tag,
        "asset_sys_id": asset_sys_id,
        "userid": userid
    }

    onboarding_details.append(onboarding_record)

    print("\n------------------------------")
    # print("Task:", task_number)
    # print("Employee:", employee_name)
    # print("Email:", email)
    # print("Location:", employee_location)
    # print("Onboarding Date:", onboarding_date)
    # print("Asset:", asset)
    # print("Asset Tag:", asset_tag)
    # print("Asset Sys ID:", asset_sys_id)
    # print("UserID:", userid)


from datetime import datetime, date, timedelta

# Calculate next Monday
today = date.today()
days_until_monday = (7 - today.weekday()) % 7

if days_until_monday == 0:
    days_until_monday = 7

next_monday = today + timedelta(days=days_until_monday)

print("Next Monday:", next_monday.strftime("%m/%d/%Y"))

# Filter employees onboarding next Monday
monday_candidates = []

for record in onboarding_details:

    onboarding_date = record["onboarding_date"]

    if not onboarding_date:
        continue

    hire_date = datetime.strptime(
        onboarding_date, "%m/%d/%Y"
    ).date()

    if hire_date == next_monday:
        monday_candidates.append(record)

# Display results
print("\nEmployees onboarding next Monday:", len(monday_candidates))
print("monday_candidates:", monday_candidates)
for candidate in monday_candidates:
    print("------------------------------")
    print("Employee:", candidate["employee_name"])
    print("Email:", candidate["email"])
    print("Location:", candidate["location"])
    print("Onboarding Date:", candidate["onboarding_date"])
    print("Asset:", candidate["asset"])
    print("Task:", candidate["task_number"])



# Get unique employees based on employee name
unique_employees = {}

for candidate in monday_candidates:
    employee_name = candidate["employee_name"]

    if employee_name not in unique_employees:
        unique_employees[employee_name] = {
            "name": employee_name,
            "location": candidate["location"],
            "email": candidate["email"],
            "userid": candidate["userid"]
        }

# Convert to a list
unique_employees_list = list(unique_employees.values())

# # # Export to JSON
# # with open("monday_onboarding_employees.json", "w", encoding="utf-8") as file:
# #     json.dump(unique_employees_list, file, indent=4, ensure_ascii=False)

# print("Unique employees:", len(unique_employees_list))
# print(json.dumps(unique_employees_list, indent=4, ensure_ascii=False))

# #  code to send email to the employee with all the details we have gathered from the above API calls

# print("\n--------------###########----------------------")


# # Image paths
# image_path = r"C:\Users\eldhoser\python-learning\OnboardingEmail_Automation\nvidia-ss1.png"
# image_qrIOS = r"C:\Users\eldhoser\python-learning\OnboardingEmail_Automation\qr-ios.png"
# image_qrAndroid = r"C:\Users\eldhoser\python-learning\OnboardingEmail_Automation\QR-android.png"

# # Start Outlook once, outside the loop
# outlook = win32.Dispatch("Outlook.Application")

# for unique_employee in unique_employees_list:

#     # Employee details
#     newjoinername = unique_employee["name"]
#     newjoinerfirstname = newjoinername.split()[0]
#     newjoineremail = unique_employee["email"]
#     newjoineruserid = unique_employee["userid"]
#     newjoinerlocation = unique_employee["location"]
#     mail = outlook.CreateItem(0)
#     emailTo = newjoineremail
#     mailsubject = f"Welcome to NVIDIA {newjoinerfirstname}!"

#     # Complete HTML email body
#     mailbody = f"""
#     <html>
#     <body style="font-family: Calibri, sans-serif; font-size: 11pt; color: #333333;">



#     <h3 style="color: #000000;">
#     Welcome to NVIDIA {newjoinerfirstname}!
#     </h3>

#     <p>
#     Below are the details to help you get set up with your laptop today.
#     </p>

#     <h3 style="color: #5c85d6; font-weight: bold; text-decoration: underline;">
#      Your Account Details
#     </h3>

#     <table style="border-collapse: collapse; width: 600px;">
#      <tr>
#          <td style="font-weight: bold; padding: 8px;">
#              Display Name:
#          </td>
#          <td style="padding: 8px;">
#              {newjoinername}
#          </td>
#      </tr>

#      <tr>
#          <td style="font-weight: bold; padding: 8px;">
#              Username:
#          </td>
#          <td style="padding: 8px;">
#              {newjoineruserid}
#          </td>
#      </tr>

#      <tr>
#          <td style="font-weight: bold; padding: 8px;">
#              Email / Username:
#          </td>
#          <td style="padding: 8px;">
#              {newjoineremail}
#          </td>
#      </tr>
#  </table>
#  <br>

#  <p style="color: red; font-weight: bold;">
#      No password is needed. You will create one.
#  </p>

#  <br>

#     <h3 style="color: #5c85d6; font-weight: bold; text-decoration: underline;">
#      Windows Deployment
#     </h3>

#  <p>
#      Deployment of NVIDIA computers is an automated process. To initiate:
#  </p>

#  <ol>
#      <li>Unbox the system and connect it to power.</li>
#      <li>Power on the machine and it will load to the Windows setup screen.</li>
#      <li>Follow the on-screen prompts to select your region and keyboard layout.</li>
#      <li>Join the system to your Wi-Fi or connect using an Ethernet cable.</li>
#      <li>
#          Once the system is connected, you will be prompted to enter your
#          setup email address. Enter your NVIDIA email address to continue.
    
#          <br>
#  <img src="cid:nvidia_logo" width="180" style="display: block;">
#  <br>
#  <li>Click <span style="font-weight: bold; color: #00cc00;">Forgot Password</span>when prompted to enter password.</li>
#  <li>When prompted - Enter code from Phone or Alternate Email , this is time-sensitive.</li>
#  <li>Follow prompts to create your first password</li>
#  <li>Complete the 2FA verification using MS Authenticator</li>
#  <li>Continue following all Windows configuration prompts</li>
#  <li>Set a “Windows Hello for Business” PIN or Face Scan.</li>
#  </li>
#  </ol>
#  <p>
#  The setup takes up to 1 hour to complete but requires no further input. Once completed you will be presented with the Windows Logon screen at which point you can log in with your Nvidia credentials.
#  </p>


#  <h3 style="color: #00cc00; font-weight: bold;">
#  Critical Updates Required!
#  </h3>
#   <P> As with all new technologies, updates are available regularly, for the best experience this process is critical: </p>
#   <ol>
#  <li>Start Menu \ Windows Updates : reboots may be required </li>
#  <li>SupportAssist or Dell Command Update app – reboots may be required </li>
#   </ol>
#   <p> Use the <span style="font-weight: bold; color: #00cc00;">Company Portal app</span> to install additional software apps and updates </p>

#  <h3 style="color: #5c85d6; font-weight: bold; text-decoration: underline;">
#     Setup Bitlocker PIN:
#  </h3>
#  <ol>
#  <li> Start menu \ type <span style="font-weight: bold; color: #00cc00;">“Manage BitLocker” </span> and open the app </li>
#  <li>Turn on BitLocker and follow the prompts to create your BitLocker PIN</li>
#  </ol>
#  <h3 style="color: #5c85d6; font-weight: bold; text-decoration: underline;">
#    Security: Set up MS Authenticator, YubiKey and Mobile PassKey:
#  </h3>

#   <p>Nvidia uses 2FA and <a href="https://nvidia.service-now.com/aiux/employeeworks/knowledge/KB0021050" target="_blank">FIDO2 Authentication</a>.  FIDO2 authentication includes Windows Hello for Business (Windows), Touch ID (Mac), YubiKeys (Linux \ MultiOS) and Mobile Passkeys) - <span style="font-weight: bold; color: #00cc00;">Register within 24 hours to avoid losing access! </span> </p>
#  <ol>
#  <li>Install and set up the <span style="font-weight: bold; color: #00cc00;"> MS Authenticator app: </span></li>
#  <table style="width: 100%; border-collapse: collapse;">
#      <tr>
#          <!-- iOS -->
#          <td style="width: 50%; vertical-align: top; padding-right: 20px;">

#              <h3 style="color: #5c85d6; font-weight: bold; text-decoration: underline;">
#                  <a href="https://apps.apple.com/gb/app/microsoft-authenticator/id983156458" target="_blank">iOS Version:</a>
#              </h3>

#              <img src="cid:qr_ios" width="180" style="display: block;">

#          </td>

#          <!-- Android -->
#          <td style="width: 50%; vertical-align: top; padding-left: 20px;">

#              <h3 style="color: #5c85d6; font-weight: bold; text-decoration: underline;">
#                  <a href="https://play.google.com/store/apps/details?id=com.microsoft.authenticator" target="_blank">Android Version:</a>
#              </h3>

#              <img src="cid:qr_android" width="180" style="display: block;">

#          </td>
#      </tr>
#  </table>

#  <li>Sign in with your NVIDIA email / password and authenticate using any listed method </li>
#  <li> Follow prompts to install and configure MS Authenticator on your mobile </li>
#  </ol>

#  <p>To set up a YubiKey (if applicable), see <a href="https://nvidia.service-now.com/aiux/employeeworks/knowledge/KB0021050" target="_blank">YubiKey Setup and Troubleshooting FAQ</a></p>
#  <p>Configure your Mobile Passkey Authentication - Passkey allows authentication to NVIDIA resources without entering a password</p>

#  <p style="color: red; font-weight: bold;">
#      Note: Mobile Authentication requires iOS 17 or Android 15
#  </p>

#  <ol>
#  <li>Launch the Authenticator app on your phone</li>
#  <li>Tap on your NVIDIA email account</li>
#  <li>Next, Tap on Create Passkey</li>
#  <li>Follow the prompts on screen and complete your PassKey setup</li>
#  <li>Congratulations! Your Passkey setup is now complete and can be used for authentication</li>
#   </ol>
#  <p>
#  For more details on FIDO authentication, see <a href="https://nvidia.service-now.com/aiux/employeeworks/knowledge/KB0021050" target="_blank">Knowledge - FIDO2 Authentication</a>
#  <br>
#  Get Help – NvidiaHelp:
#  <br>
#  Any IT problems or service requests should be logged via the NvidiaHelp system:

#  <ol>
#  <li>Online portal \ Chat \ Knowledgebase:<a href="https://nvidia.service-now.com/aiux/employeeworks/knowledge/KB0021050" target="_blank"> NVIDIA Help</a></li>
#  <li>Email: NVIDIAHelp@nvidia.com</li>
#  <li>Tel: +49 2405 478001 (Germany)</li>
#  </ol>

#  </p>
#  </body>
#  </html>
#  """

#     # -----------------------------
#     # Create Outlook email
#     # -----------------------------

#     outlook = win32.Dispatch("Outlook.Application")
#     email = outlook.CreateItem(0)
#     email.To = emailTo
#     email.Subject = mailsubject
#     email.HTMLBody = mailbody

#     # -----------------------------
#     # Add NVIDIA logo
#     # -----------------------------

#     attachment = email.Attachments.Add(
#     image_path,
#     1,                  # olByValue
#     0,                  # Position
#     "nvidia_logo.png"
#     )

#     attachment.PropertyAccessor.SetProperty(
#     "http://schemas.microsoft.com/mapi/proptag/0x370E001E",
#     "image/png"
#     )

#     attachment.PropertyAccessor.SetProperty(
#     "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
#     "nvidia_logo"
#     )


#     # -----------------------------
#     # Add iOS image
#     # -----------------------------

#     ios_attachment = email.Attachments.Add(
#     image_qrIOS,
#     1,                  # olByValue
#     0,
#     "qr_ios.png"
#     )

#     ios_attachment.PropertyAccessor.SetProperty(
#     "http://schemas.microsoft.com/mapi/proptag/0x370E001E",
#     "image/png"
#     )

#     ios_attachment.PropertyAccessor.SetProperty(
#     "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
#     "qr_ios"
#     )


#     # -----------------------------
#     # Add Android image
#     # -----------------------------

#     android_attachment = email.Attachments.Add(
#     image_qrAndroid,
#     1,
#     0,
#     "qr_android.png"
#     )

#     android_attachment.PropertyAccessor.SetProperty(
#     "http://schemas.microsoft.com/mapi/proptag/0x370E001E",
#     "image/png"
#     )

#     android_attachment.PropertyAccessor.SetProperty(
#     "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
#     "qr_android"
#     )


#     # -----------------------------
#     # Set HTML body
#     # -----------------------------
#     email.HTMLBody = mailbody

#     # Display each email as a draft for review
#     email.Display()

#     print(f"Draft created for {newjoinername}: {newjoineremail}")