
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
    asset = task.get("cmdb_ci", {}).get("display_value", "")
    

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
###########

response = requests.get(
    f"{base_url}/alm_asset",
    params={
        "sysparm_query": f"assigned_to={user_sys_id}",
        "sysparm_fields": "sys_id,asset_tag,display_name,model,assigned_to,serial_number",
        "sysparm_display_value": "all"
    },
    headers=headers,
    auth=(client_id, client_secret),
    timeout=(10, 120)
)

response.raise_for_status()

assets = response.json().get("result", [])

for item in assets:
    #print(item)
    laptop_asset = ""

laptop_asset = ""
asset_tag = ""
laptop_sys_id = ""

for item in assets:
    model = item.get("model", {})
    model_name = model.get("display_value", "") if isinstance(model, dict) else model

    display = item.get("display_name", {})
    display_name = display.get("display_value", "") if isinstance(display, dict) else display

    if "backpack" not in model_name.lower() and "logitech" not in model_name.lower():

        laptop_asset = display_name

        sys_id = item.get("sys_id", {})
        laptop_sys_id = sys_id.get("value", "") if isinstance(sys_id, dict) else sys_id

        tag = item.get("asset_tag", {})
        asset_tag = tag.get("display_value", "") if isinstance(tag, dict) else tag

        print("Selected Asset:", laptop_asset)
        print("Asset Tag:", asset_tag)
        print("Asset Sys ID:", laptop_sys_id)
        break

##########

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
        "userid": userid
    }

    onboarding_details.append(onboarding_record)

    print("\n------------------------------")
    print("Task:", task_number)
    print("Employee:", employee_name)
    print("Email:", email)
    print("Location:", employee_location)
    print("Onboarding Date:", onboarding_date)
    print("Asset:", asset)
    print("UserID:", userid)

response = requests.get(
    f"{base_url}/alm_asset",
    params={
        "sysparm_query": f"assigned_to={user_sys_id}",
        "sysparm_fields": "sys_id,asset_tag,display_name,model,assigned_to,serial_number",
        "sysparm_display_value": "all"
    },
    headers=headers,
    auth=(client_id, client_secret),
    timeout=(10, 120)
)

response.raise_for_status()

assets = response.json().get("result", [])

for item in assets:
    print(item)