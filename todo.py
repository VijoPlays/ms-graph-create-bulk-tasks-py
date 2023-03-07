# For more info read: https://learn.microsoft.com/en-us/graph/api/resources/planner-overview?view=graph-rest-1.0

import json

import requests


baseURL = "https://graph.microsoft.com/v1.0/me/todo/lists/"
tokenFile = open(".token", "r")
authToken = tokenFile.read()
headers = {
    "Content-Type": "application/json",
    "Authorization": "Bearer " + authToken,
}

# Set these variables the way you need them.
taskCount = 1000
listCount = 1


# Creates a List in Todo and returns the ID of it.
#
# 	Count is simply some number, used for giving the bulkList an identifier. For single plans, just supply "1".
def createTodoList(count):
    body = json.dumps({"displayName": "bulkList_" + str(count)})

    r = requests.post(url=baseURL, data=body, headers=headers)
    data = r.json()

    return data["id"]


# Creates a Task in a list.
#
# 	ListID is the list it should be assigned to.
# 	Count is simply some number, used for giving the bulkTask an identifier. For single tasks, just supply "1".
def createTodoTask(listID, count):
    url = baseURL + listID + "/tasks"
    body = json.dumps({"title": "bulkTask_" + str(count)})

    r = requests.post(url=url, data=body, headers=headers)
    data = r.json()

    return data["id"]


# Deletes a list with the provided listID.
def deleteTodoList(listID):
    url = baseURL + str(listID)
    headerDelete = {
        "Authorization": "Bearer " + authToken,
    }

    r = requests.delete(url=url, headers=headerDelete)

    if r.status_code == 204:
        print("successfully deleted the plan")
    else:
        print(
            "error occurred upon deletion of plan. status code is: "
            + str(r.status_code)
        )


listIDs = []

i = 0
j = 0
while i < listCount:
    listID = createTodoList(count=i)
    print("ListID for deletion is: " + listID)
    while j < taskCount:
        createTodoTask(listID=listID, count=j)
        print("Created task no.: " + str(j) + " in list no.: " + str(i))
        j += 1
    j = 0
    listIDs.append(listID)
    i += 1

# for listID in listIDs:
#     deleteTodoList(listID=listID)

print(
    "If an error occurred, it's most likely related to a missing token. Check the README for more details."
)
