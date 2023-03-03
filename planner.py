import json
import re
import requests

# https://learn.microsoft.com/en-us/graph/api/resources/planner-overview?view=graph-rest-1.0

baseURL = "https://graph.microsoft.com/v1.0/"
tokenFile = open(".token", "r")
authToken = tokenFile.read()
headers = {
    "Content-Type": "application/json",
    "Authorization": "Bearer " + authToken,
}

# Set these variables the way you need them.
taskCount = 300


# Gets an array of groupIDs from Planner.
def getPlannerGroups():
    url = baseURL + "me/getMemberGroups"
    body = json.dumps({"securityEnabledOnly": False})

    r = requests.post(url=url, data=body, headers=headers)
    return r.json()["value"]


# Creates a Plan in Planner and returns the ID of it.
def createPlannerPlan(groupID):
    url = baseURL + "planner/plans"
    body = json.dumps(
        {
            "container": {  # TODO: Replace the groupID below this one
                "url": "https://graph.microsoft.com/v1.0/groups/" + groupID
            },
            "title": "bulkPlan",
        }
    )

    r = requests.post(url=url, data=body, headers=headers)
    data = r.json()

    return data["id"]


# Creates a Task in a plan.
#   PlanID is the plan it should be assigned to.
#   Count is simply some number, used for giving the bulkTask an identifier. For single tasks, just supply "1".
def createPlannerTask(planID, count):
    url = baseURL + "planner/tasks"
    body = json.dumps({"planId": planID, "title": "bulkTask_" + str(count)})

    r = requests.post(url=url, data=body, headers=headers)
    data = r.json()

    return data["id"]


# FIXME: Couldn't get deletion to work, have to delete the plan itself manually.
# def getPlannerPlanDetails(planID):
#     url = baseURL + "planner/plans/" + planID + "/details"
#     r = requests.get(url=url, headers=headers)
#     data = r.json()

#     eTag = data["@odata.etag"].removeprefix("W/")
#     print(eTag)
#     return eTag


# def deletePlannerPlan(planID):
#     eTag = getPlannerPlanDetails(planID=planID)
#     url = baseURL + "planner/plans/" + planID
#     headerDelete = {
#         "Authorization": "Bearer " + authToken,
#         "If-Match": eTag,
#     }

#     r = requests.delete(url=url, headers=headerDelete)

#     print(r.json())
#     print("Status code for deletion was: " + str(r.status_code))

groupIDs = getPlannerGroups()
planID = createPlannerPlan(groupIDs[0])
print("PlanID for deletion is: " + planID)

i = 0
while i < taskCount:
    createPlannerTask(planID, i)
    print("Created task no.: " + str(i))
    i += 1
