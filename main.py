# -*- coding: UTF-8 -*-
import json
import time
import sys
import requests as req

path = sys.path[0] + r'/rtk.txt'
# path = r'./rtk.txt'




def gettoken(refresh_token):
    headers = {'Content-Type': 'application/x-www-form-urlencoded'
               }
    data = {'grant_type': 'refresh_token',
            'refresh_token': refresh_token,
            'client_id': id,
            'client_secret': secret,
            'redirect_uri': 'http://localhost:53682/'
            }
    html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', data=data, headers=headers)
    jsontxt = json.loads(html.text)
    refresh_token = jsontxt['refresh_token']
    access_token = jsontxt['access_token']
    with open(path, 'w+') as f:
        f.write(refresh_token)
    return access_token


def main():
    fo = open(path, "r+")
    refresh_token = fo.read()
    fo.close()
    localtime = time.asctime(time.localtime(time.time()))
    access_token = gettoken(refresh_token)
    headers = {
        'Authorization': access_token,
        'Content-Type': 'application/json'
    }

    print('获取app信息')
    req.get(r'https://api.powerbi.com/v1.0/myorg/apps', headers=headers)

    print('正在获取团队和用户信息')
    me = req.get(r'https://graph.microsoft.com/v1.0/me', headers=headers).text
    me_json = json.loads(me)
    my_mail = me_json['mail']
    # print('me:', me)
    users = req.get(r'https://graph.microsoft.com/v1.0/users', headers=headers).text
    # print('users:', users)
    groups = req.get(r'https://graph.microsoft.com/v1.0/groups', headers=headers).text
    # print('groups:', groups)
    groups_json = json.loads(groups)
    group_id = groups_json['value'][0]['id']
    # print('group_id:', group_id)

    print("正在确认日历")
    req.get(r'https://graph.microsoft.com/v1.0/me/calendar', headers=headers)
    time.sleep(1)

    print("正在获取活动信息")
    events = req.get(r'https://graph.microsoft.com/v1.0/groups/' + group_id + '/calendar/events',
                     headers=headers).text
    # print('events:', events)

    print('正在安排活动')
    new_event = {
        "subject": "Union Subscription Update " + localtime,
        "body": {
            "contentType": "HTML",
            "content": "Running a cloud-based union subscription update service." + localtime
        },
        "start": {
            "dateTime": "2022-07-2T12:00:00",
            "timeZone": "Pacific Standard Time"
        },
        "end": {
            "dateTime": "2022-07-2T12:10:00",
            "timeZone": "Pacific Standard Time"
        }
    }
    event = req.post(r'https://graph.microsoft.com/v1.0/groups/' + group_id + '/calendar/events', headers=headers,
                     json=new_event).text
    event_json = json.loads(event)
    event_id = event_json['id']

    print('正在确认活动')
    events = req.get(r'https://graph.microsoft.com/v1.0/groups/' + group_id + '/calendar/events',
                     headers=headers).text
    # print('events:', events)
    time.sleep(1)

    print('正在移除活动')
    req.delete(r'https://graph.microsoft.com/v1.0/groups/' + group_id + '/calendar/events/' + event_id,
               headers=headers)

    print('确认移除活动')
    events = req.get(r'https://graph.microsoft.com/v1.0/groups/' + group_id + '/calendar/events',
                     headers=headers).text
    # print('events:', events)
    time.sleep(1)

    print('正在获取代办事项')
    todo_task_lists = req.get(r'https://graph.microsoft.com/v1.0/me/todo/lists',
                              headers=headers).text
    # print('todo_task_lists:', todo_task_lists)
    todo_task_lists_json = json.loads(todo_task_lists)
    todo_task_list_id = todo_task_lists_json['value'][0]['id']
    todo_tasks = req.get(r'https://graph.microsoft.com/v1.0/me/todo/lists/' + todo_task_list_id + '/tasks',
                         headers=headers).text
    # print('todo_tasks:', todo_tasks)

    print('创建代办事项')
    new_task = {
        "title": "Syndicated Subscription Update " + localtime,
        "categories": ["Important"],
        "linkedResources": [
            {
                "webUrl": "http://microsoft.com",
                "applicationName": "Microsoft",
                "displayName": "Microsoft"
            }
        ]
    }
    task = req.post(r'https://graph.microsoft.com/v1.0/me/todo/lists/' + todo_task_list_id + '/tasks',
                    headers=headers, json=new_task).text
    task_json = json.loads(task)
    task_id = task_json['id']

    print('正在确认代办事项')
    todo_tasks = req.get(r'https://graph.microsoft.com/v1.0/me/todo/lists/' + todo_task_list_id + '/tasks',
                         headers=headers).text
    # print('todo_tasks:', todo_tasks)
    time.sleep(1)

    print('正在移除待办事项')
    req.delete(r'https://graph.microsoft.com/v1.0/me/todo/lists/' + todo_task_list_id + '/tasks/' + task_id,
               headers=headers)

    print('确认移除代办事项')
    todo_tasks = req.get(r'https://graph.microsoft.com/v1.0/me/todo/lists/' + todo_task_list_id + '/tasks',
                         headers=headers).text
    # print('todo_tasks:', todo_tasks)

    print('正在确认驱动器信息')
    req.get(r'https://graph.microsoft.com/v1.0/me/drive', headers=headers)
    time.sleep(1)
    req.get(r'https://graph.microsoft.com/v1.0/me/drive/root', headers=headers)
    time.sleep(1)
    req.get(r'https://graph.microsoft.com/v1.0/me/drive/root/children', headers=headers)
    time.sleep(1)
    req.get(r'https://graph.microsoft.com/v1.0/me/drive/root/versions', headers=headers)
    time.sleep(1)

    print('邮件检查并整理')
    req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders', headers=headers)
    time.sleep(1)
    req.get(r'https://graph.microsoft.com/v1.0/me/outlook/masterCategories', headers=headers)
    time.sleep(1)
    req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules', headers=headers)
    time.sleep(1)
    messages = req.get(r'https://graph.microsoft.com/v1.0/me/messages', headers=headers).text
    # print('messages:', messages)
    messages_json = json.loads(messages)
    for message in messages_json['value']:
        message_id = message['id']
        rs = req.delete(f'https://graph.microsoft.com/v1.0/me/messages/{message_id}', headers=headers)
        # print('rs:', rs.status_code)
    time.sleep(1)

    print('联合更新邮件推送')
    new_mail = {
        "message": {
            "subject": "联合更新邮件推送" + localtime,
            "body": {
                "contentType": "Text",
                "content": "Completing a cloud-based union subscription update service."
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": my_mail
                    }
                }
            ]
        },
        "saveToSentItems": "false"
    }
    req.post(r'https://graph.microsoft.com/v1.0/me/sendMail', headers=headers, json=new_mail)

    print("联合订阅更新完成")


if __name__ == '__main__':
    main()
