import argparse
import json
import os
import time
import requests

# For more info read: https://learn.microsoft.com/en-us/graph/api/resources/planner-overview?view=graph-rest-1.0

baseURL = "https://graph.microsoft.com/v1.0/"
tokenFile = open(".token", "r")
authToken = tokenFile.read()
tokenFile.close()
headers = {
    "Content-Type": "application/json",
    "Authorization": "Bearer " + authToken,
}
plannerPlanFileName = "store" + os.sep + "plannerPlans.txt"

parser = argparse.ArgumentParser(
    description="By adding a -d flag you can delete all created plans."
)
parser.add_argument("-d", action="store_true")
providedCLFlags = parser.parse_args()

# Set these variables the way you need them.
taskCount = 10000
planCount = 1


# Gets an array of groupIDs from Planner.
def getPlannerGroups():
    url = baseURL + "me/getMemberGroups"
    body = json.dumps({"securityEnabledOnly": False})

    r = requests.post(url=url, data=body, headers=headers)
    return r.json()["value"]


# Creates a Plan in Planner and returns the ID of it.
#   GroupID is the plan it should be assigned to.
#   Count is simply some number, used for giving the bulkTask an identifier. For single tasks, just supply "1".
def createPlannerPlan(groupID, count):
    url = baseURL + "planner/plans"
    body = json.dumps(
        {
            "container": {"url": "https://graph.microsoft.com/v1.0/groups/" + groupID},
            "title": "bulkPlan_" + str(count),
        }
    )

    r = requests.post(url=url, data=body, headers=headers)
    data = r.json()

    id = data["id"]

    file = open(plannerPlanFileName, "a")
    file.write(id + "\n")
    file.close()

    return id


# Creates a Task in a plan.
#   PlanID is the plan it should be assigned to.
#   Count is simply some number, used for giving the bulkTask an identifier. For single tasks, just supply "1".
def createPlannerTask(planID, count):
    url = baseURL + "planner/tasks"
    body = json.dumps({"planId": planID, "title": "bulkTask_" + str(count)})

    r = requests.post(url=url, data=body, headers=headers)
    data = r.json()

    if r.status_code == 429:
        t = r.headers.get("retry-after")
        time.sleep(int(t))
        createPlannerTask(planID=planID, count=count)
        return

    return data["id"]


# Retrieves an eTag from a plan, if the plan exists on the group.
def getPlannerPlanETag(groupID, planID):
    url = baseURL + "groups/" + groupID + "/planner/plans"
    r = requests.get(url=url, headers=headers)
    data = r.json()

    plans = data["value"]
    for plan in plans:
        if plan["id"] == planID:
            return plan["@odata.etag"]

    print("ERROR: Could not find plan " + planID + " in group " + groupID + ".")
    exit


# Deletes a plan from the plannerPlans.txt file in group of groupID.
def deletePlannerPlan(planID, groupID):
    eTag = getPlannerPlanETag(planID=planID, groupID=groupID)
    url = baseURL + "planner/plans/" + planID
    headerDelete = {
        "Authorization": "Bearer " + authToken,
        "If-Match": eTag,
    }

    r = requests.delete(url=url, headers=headerDelete)

    if r.status_code == 204:
        print("successfully deleted the plan with ID: " + planID)
    else:
        print(
            "error occurred upon deletion of plan. status code is: "
            + str(r.status_code)
        )


groupIDs = getPlannerGroups()
groupID = groupIDs[0]
print("GroupID for deletion is: " + groupID)

if providedCLFlags.d:
    # Delete plan routine.
    planIDs = []
    with open(plannerPlanFileName) as file:
        for line in file:
            planIDs.append(line.removesuffix("\n"))
    open(plannerPlanFileName, "w").close()

    for planID in planIDs:
        deletePlannerPlan(planID=planID, groupID=groupID)
else:
    # Create plan routine.
    i = 0
    j = 0
    while i < planCount:
        planID = createPlannerPlan(groupID=groupID, count=i)
        print("PlanID for deletion is: " + planID)
        while j < taskCount:
            createPlannerTask(planID=planID, count=j)
            print("Created task no.: " + str(j) + " in plan no.: " + str(i))
            j += 1
        j = 0
        i += 1

print(
    "If an error occurred, it's most likely related to a missing token. Check the README for more details."
)
