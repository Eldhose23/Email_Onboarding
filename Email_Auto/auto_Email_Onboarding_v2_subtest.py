
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

    response.raise_for_status()
    tasks = response.json().get("result", [])

    if not tasks:
        break

    all_tasks.extend(tasks)

    print(f"Retrieved {len(tasks)} tasks (total: {len(all_tasks)})")

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
            "sysparm_fields": "number,u_task_requested_for",
            "sysparm_display_value": "all"
        },
        headers=headers,
        auth=(client_id, client_secret),
        timeout=(10, 120)
    )

    response.raise_for_status()
    task_details = response.json().get("result", {})

    user_sys_id = (
        task_details.get("u_task_requested_for", {}).get("display_value", "")
    )
    sctask= (
            task_details.get("number", {}).get("value", "")
        )
    print(user_sys_id,sctask)
    
    if not user_sys_id:
        print(f"Skipping {task_number}: no requested employee found")
        continue
   # print(f"\nTotal IDs: {len(user_sys_id)}")


response = requests.get(
    f"{base_url}/sys_user/68af327ec3d4d71043b55d0fb001317a",
    params={
        "sysparm_fields": "sys_id,name,user_name,email,active,location",
        "sysparm_display_value": "all"
    },
    headers=headers,
    auth=(client_id, client_secret),
    timeout=(10, 120)
)

print("HTTP status:", response.status_code)
print("Response body:", response.text)


 # Get onboarding date records for this employee
response = requests.get(
        f"{base_url}/u_onboarding_changes",
        params={
            "sysparm_query": (
                f"u_user_profile=68af327ec3d4d71043b55d0fb001317a"
                 "^ORDERBYDESCsys_updated_on"
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

print("HTTP status:", response.status_code)
print("Response body:", response.text)


  #########


response.raise_for_status()
records = response.json().get("result", [])