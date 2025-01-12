import firebase_admin
from firebase_admin import storage, firestore,credentials
import firebase_admin.messaging
from firebase_functions import https_fn, storage_fn

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Side, Border, Font
from openpyxl.styles.colors import Color
from io import BytesIO
import time
import uuid
import datetime

from firebase_functions.firestore_fn import (
  on_document_written,
  Event,
  DocumentSnapshot,
  Change 
)


cred = credentials.Certificate('shiftmaker-fee68-firebase-adminsdk-9c2ic-b503fe46e1.json')
firebase_admin.initialize_app(cred, {
    'storageBucket': 'gs://shiftmaker-fee68.appspot.com'
})

db = firestore.client()

@https_fn.on_request(region="asia-northeast1")
def make_directory(req: https_fn.Request) -> https_fn.Response:
    request_json = req.get_json(silent=True)
    userId = request_json['userId']
    fullname = request_json['fullname']
    time.sleep(2)
    print("request is arrived")
    print(userId)
    print(fullname)
    db.collection("Users").document(userId).collection("MyInfo").document("userInfo").set({"username": fullname,"groupId":{"1":"no data"},"birthday":"2000-01-01","hourlyWage":{"no data":1000},"tokens":"no data"})
    db.collection("Users").document(userId).collection("MyInfo").document("currentGroupNum").set({"currentNum":"1"})
    db.collection("Users").document(userId).collection("MyInfo").document("groups").set({"1":{"groupName":"no data","groupId":"no data"}})
    db.collection("Users").document(userId).collection("RequestShift").document("shift").set({"no data":{"開始": "no data", "終了": "no data"}})
    db.collection("Users").document(userId).collection("CompletedShift").document("shift").set({"no data":{"2000/01/01":"12:00 - 12:00"}})
    
    return https_fn.Response("ok")


@https_fn.on_call(region="asia-northeast1")
def create_group(req: https_fn.CallableRequest):
    print("request is arrived")
    print(req)
    userId = req.data["userId"]
    groupName =req.data["groupName"]
    groupPass =req.data["pass"]
    groupId =uuid.uuid4()
    print(type(groupId))
    print(type(str(groupId)))
    groupId = str(groupId)
    print(f"userId:{userId} groupName:{groupName} grouPass:{groupPass} groupId:{groupId}")

    adminname=db.collection("Users").document(userId).collection("MyInfo").document("userInfo").get().to_dict()["username"]
    group_ref =db.collection("Users").document(userId).collection("MyInfo").document("groups")
    group_data =group_ref.get().to_dict()
    firstGroup = group_data["1"]

    print(f"group_data:{group_data}")
    print(f"group_data[groupId]:{firstGroup["groupId"]}")
    print(firstGroup["groupId"]=="no data")
    

    if firstGroup["groupId"]=="no data":
      groupCount="1"
    else:
      groupCount=str(len(group_data)+1)
      print(groupCount)
    group_ref.update({groupCount:{"groupName":groupName,"groupPass":groupPass,"groupId":groupId}})
      
    doc_ref_groupInfo = db.collection('Groups').document(groupId).collection("groupInfo")

    doc_ref_groupInfo.document("pass").set({"pass": groupPass,"groupName":groupName,"adminName":adminname})
    doc_ref_groupInfo.document("download").set({"downloadLink": "no data"})
    doc_ref_groupInfo.document("status").set({"status":False})
    doc_ref_groupInfo.document("admin").set({"adminId":userId})
    doc_ref_groupInfo.document("fileName").set({"file_name":"no data","comFileName":"no data"})
    doc_ref_groupInfo.document("member").set({"1": "no data"})
    doc_ref_groupInfo.document("tableRequest").set({"start":"no data","end":"no data"})
    doc_ref_groupInfo.document("applicants").set({"1":"no data"})
    doc_ref_groupInfo.document("message").set({"message":""})
    return https_fn.Response("ok")


@https_fn.on_call(region="asia-northeast1")
def delete_group(req: https_fn.CallableRequest):
  userId = req.data["userId"]
  groupId =req.data["groupId"]
  selectedNumber =int(req.data["num"])
  print(f"userId: {userId}, groupId: {groupId},selectedNumber:{selectedNumber}")
  doc_ref_delete = db.collection("Groups").document(groupId).collection("groupInfo")
  MyInfo_ref=db.collection("Users").document(userId).collection("MyInfo")
  groups=MyInfo_ref.document("groups").get().to_dict()
  new_data = {} 
  for num, data in groups.items():
    if int(num) != selectedNumber:
        print(f" {num}  {data}")
        if selectedNumber ==1:
            num = str(int(num) - 1)
            new_data[num]=data
        elif selectedNumber ==2:
            if num =="3":
                num = str(int(num) - 1)
            new_data[num]=data
        elif selectedNumber ==3:
            new_data[num]=data

  if len(new_data)==0:
      new_data["1"]={"groupId":"no data","groupName":"no data"}
  MyInfo_ref.document("currentGroupNum").update({"currentNum":"1"})
  MyInfo_ref.document("groups").set(new_data)
  if groupId !="no data":
    doc_ref_delete.document("admin").delete()
    doc_ref_delete.document("applicants").delete()
    doc_ref_delete.document("download").delete()
    doc_ref_delete.document("fileName").delete()
    doc_ref_delete.document("member").delete()
    doc_ref_delete.document("message").delete()
    doc_ref_delete.document("pass").delete()
    doc_ref_delete.document("status").delete()
    doc_ref_delete.document("tableRequest").delete()
    db.collection("Groups").document(groupId).delete()
  return https_fn.Response("ok")

@https_fn.on_call(region="asia-northeast1")
def request_group(req: https_fn.Request) -> https_fn.Response:
    request_json = req.get_json(silent=True)
    if request_json and 'userId' in request_json:
        userId = request_json['userId']
        groupId = request_json['groupId']
    
    userInfo_ref = db.collection("Users").document(userId).collection("MyInfo").document("userInfo")
    username =userInfo_ref.get().to_dict()["username"]
    
    applicants_ref = db.collection("Groups").document(groupId).collection("groupInfo").document("applicants")
    applicants_data=applicants_ref.get().to_dict()
    if applicants_data["1"] == "no data":
      applicants_count=str(len(applicants_data))
      print("no data")
      applicants_ref.update({applicants_count:{"name":username, "situation":"not yet", "uid":userId}})
    else:   
      applicants_count=str(len(applicants_data)+1)
      print(f"already :{applicants_count}")
      applicants_ref.update({applicants_count:{"name":username, "situation":"not yet", "uid":userId}})
    return https_fn.Response("ok")



@https_fn.on_call(region="asia-northeast1")
def admit_member(req: https_fn.Request) -> https_fn.Response:
  print(f"req:{req.data}")
  userId = req.data['userId']
  groupId = req.data['groupId']
  groupName = req.data['groupName']
  print(userId)
  print(groupId)
  print(groupName)

  userInfo_ref = db.collection("Users").document(userId).collection("MyInfo").document("userInfo")
  completedShift_ref = db.collection("Users").document(userId).collection("CompletedShift").document("shift")
  userGroupIdData = userInfo_ref.get().to_dict()["groupId"]
  groupIdCount = len(userGroupIdData)
  userHourlyWage = userInfo_ref.get().to_dict()["hourlyWage"]
  
  if groupIdCount == 1 and userGroupIdData["1"] == "no data":
    userGroupIdData["1"] = groupId
  else:
    userGroupIdData[f"{groupIdCount+1}"] = groupId
  
  switch = True
  i = 1 
  while switch:
    if groupName in userHourlyWage:
      groupName = groupName + f"({i})"
      i = i + 1
    else:
      switch = False
          
  userHourlyWage[f"{groupName}"] = 1000
  completedShift_ref.update({f"{groupName}":{"2000/01/01":"12:00 - 12:00"}})
  userInfo_ref.update({"groupId":userGroupIdData,"hourlyWage":userHourlyWage})



@https_fn.on_call(region="asia-northeast1")
def delete_member(req: https_fn.Request) -> https_fn.Response:
  request_json = req.get_json(silent=True) 
  userId = request_json['userId']
  groupId = request_json['groupId']
  
  userInfo_ref = db.collection("Users").document(userId).collection("MyInfo").document("userInfo")
  userGroupIdData = userInfo_ref.get().to_dict()["groupId"]
  groupIdCount = len(userGroupIdData)

  print(f"userGroupIdData : \n {userGroupIdData}")
  print(f"groupIdCount : \n {groupIdCount}")
  
  newUserGroupIdData = {}
  borderSwitch = False
  print(groupId in userGroupIdData.values())
  if groupId in userGroupIdData.values():
    for i in range(1,groupIdCount+1,1):
        print(userGroupIdData[f"{i}"])
        if userGroupIdData[f"{i}"] != groupId:
          if borderSwitch == False:
            newUserGroupIdData[f"{i}"] =userGroupIdData[f"{i}"]
          if borderSwitch == True:
            newUserGroupIdData[f"{i-1}"] =userGroupIdData[f"{i}"]               
        else:
          borderSwitch = True
          print("ooo")
    if newUserGroupIdData.get("1") is None:
      print("bbb")
      newUserGroupIdData["1"] = "no data"
    userInfo_ref.update({"groupId":newUserGroupIdData})
  print("done")





@on_document_written(document="Groups/{groupId}/groupInfo/tableRequest",
                     region="asia-northeast1")
def generate_excel_file(
    event: Event[Change[DocumentSnapshot]],
  ) -> None:
  #ユーザーIDとタイトルを取得
  groupId = event.params["groupId"]
  tableInfo = event.data.after.to_dict()
  table_title = tableInfo['tableTitle']
  member =db.collection("Groups").document(groupId).collection("groupInfo").document("member").get().to_dict()
  members = {}
  for i in member:
   members[i] = member[f"{i}"]["name"]  

  YOUBI = ("月", "火", "水", "木", "金", "土", "日")

  # storageへの参照を取得
  bucket = storage.bucket("shiftmaker-fee68.appspot.com")

  # Workbookオブジェクトの作成-----------------------------
  wb = Workbook()
  ws = wb.create_sheet("worksheet", 0)
  #計算----------------------------------------------------
  #日数計算
  start = tableInfo["start"].split("-")
  end = tableInfo["end"].split("-") 
  d1 = datetime.date(year=int(start[0]), month=int(start[1]), day=int(start[2]))
  d2 = datetime.date(year=int(end[0]), month=int(end[1]), day=int(end[2]))
  td = abs(d2-d1) + datetime.timedelta(days=1)
  #人数とセル結合後の枠の計算
  rows=len(members)*2
  # 書式設定の変更-----------------------------------------
  for row in ws[f'A1:{ws.cell(row=rows+3,column=td.days+2).coordinate}']:
      for cell in row:
          cell.alignment = Alignment(horizontal="center", vertical="center",wrapText=True)

  for row in ws[f'C4:{ws.cell(row=rows+3,column=td.days+2).coordinate}']:
      for cell in row:
          cell.number_format ="hh:mm"

  #レイアウト作り------------------------------------------
  #罫線を引く
  border1= Border(
                                          left = Side(style="thin", color = Color("7d8591")),
                                          right = Side(style="thin", color = Color("7d8591")),
                                          top = Side(style="thin", color = Color("7d8591")),
                                          bottom = Side(style="thin", color = Color("7d8591"))                                        
                                          )
  border2 = Border(
                                          left = Side(style="thin", color = Color("000000")),
                                          right = Side(style="thin", color = Color("000000")),
                                          top = Side(style="double", color = Color("000000")),
                                          bottom = Side(style="thin", color = Color("000000"))                                        
                                          )
  border3 = Border(
                                          left = Side(style="thin", color = Color("000000")),
                                          right = Side(style="thin", color = Color("000000")),
                                          top = Side(style="thin", color = Color("000000")),
                                          bottom = Side(style="thin", color = Color("000000"))                                        
                                          )

  for row in ws[f'A2:{ws.cell(row=rows+3,column=td.days+2).coordinate}']:
    for cell in row:
        cell.border = border1
        if cell.row == 4 and cell.col_idx >=3:
          cell.border = border2
        elif cell.row >= 4 and cell.col_idx <= 2:
          cell.border = border3 

  #幅を整える・セルを結合する------------------------------------
  ws.row_dimensions[1].height = 40
  ws.row_dimensions[2].height = 25
  ws.row_dimensions[3].height = 25
  ws.column_dimensions["A"].height = 10
  ws.column_dimensions["B"].height = 10

  for num in range(4,rows+4,1):
    ws.row_dimensions[num].height = 27
    if num%2 == 0:
      cell_num=ws.cell(row=num,column=1).coordinate
      next_cell_num=ws.cell(row=num+1,column=1).coordinate
      ws.merge_cells(f"{cell_num}:{next_cell_num}")

  for num in range(3,td.days+4,1):
    column_num=ws.cell(row=1,column=num).column_letter  
    ws.column_dimensions[column_num].width = 7

  #要素を入力(シフト・氏名等)--------------------------------
  ws.cell(row=1,column=1).value =f"{start[0]}/{start[1]}/{start[2]}"
  ws.cell(row=1,column=2).value =f"{end[0]}/{end[1]}/{end[2]}"
  ws.cell(row=2,column=2).value = "日付"
  ws.cell(row=3,column=1).value = "氏名"
  ws.cell(row=3,column=2).value = "曜日"

  
  for num in range(4,rows+4,2):
    ws.cell(row=num,column=1).value = members[str(int((num/2)-1))]
    ws.cell(row=num,column=2).value = "開始"
    ws.cell(row=num+1,column=2).value = "終了"

  for num in range(3,td.days+3,1):
    # 日
    ws.cell(row=2, column=num).value = d1.day
    # 曜日
    ws.cell(row=3, column=num).value = YOUBI[d1.weekday()]
    d1 += datetime.timedelta(days=1)


  print(f'sheet name: {ws.title}')


  save_data = BytesIO()
  wb.save(save_data)
  modified_file_data = save_data.getvalue()
  dt_now = datetime.datetime.now()
  dt_now=dt_now.strftime('%Y%m%d%H%M%S')

  file_name = db.collection("Groups").document(groupId).collection("groupInfo").document("fileName").get().to_dict()["file_name"]
  if file_name != "no data":
    old_blob = bucket.blob(f'Groups/{groupId}/{file_name}')
    old_blob.delete()

  file_name = f'{table_title}{dt_now}.xlsx' 
  new_blob = bucket.blob(f'Groups/{groupId}/{file_name}')
  new_blob.upload_from_string(modified_file_data)
  update_ref = db.collection("Groups").document(groupId).collection("groupInfo").document("fileName")
  update_ref.update({"file_name": file_name})
  db.collection("Groups").document(groupId).collection("groupInfo").document("status").update({"status":True})

  membersData=db.collection("Groups").document(groupId).collection("groupInfo").document("member").get().to_dict()

  for key, value in membersData.items():
    db.collection("Groups").document(groupId).collection("groupInfo").document("member").update(
      {f"{key}":{
        "name":f"{value["name"]}",
        "situation":"not yet",
        "uid":f"{value["uid"]}"}
      })
  groupName = db.collection("Groups").document(groupId).collection("groupInfo").document("pass").get().to_dict()["groupName"]
  tokens = []
  member = db.collection("Groups").document(groupId).collection("groupInfo").document("member").get().to_dict()
  for i in range(1,len(member)+1,1):
    uid =member[f"{i}"]["uid"]
    token = db.collection("Users").document(uid).collection("MyInfo").document("userInfo").get().to_dict()["token"]
    print(token)
    tokens.append(token)
    print(f"{tokens}")
    
  message = firebase_admin.messaging.MulticastMessage(
      
      tokens=tokens,
      android=firebase_admin.messaging.AndroidConfig(
        notification=firebase_admin.messaging.AndroidNotification(
          title=f'{groupName} からアナウンス',
          body='シフトが募集されています',
        ),
        priority="high",
        direct_boot_ok=True
      ),
  )


@on_document_written(document="Users/{userId}/RequestShift/shift")
def input_data(
    event: Event[Change[DocumentSnapshot]],
  ) -> None:


  userId = event.params["userId"]
  docData = event.data.after.to_dict()
  groupId = docData["RequestedGroupId"]
  request = docData[groupId]
  userInfo = db.collection("Users").document(userId).collection("MyInfo").document("userInfo").get().to_dict()
  username = userInfo["username"]
  file_name = db.collection("Groups").document(groupId).collection("groupInfo").document("fileName").get().to_dict()["file_name"]
  print(file_name)
  status = db.collection("Groups").document(groupId).collection('groupInfo').document("status").get().to_dict()["status"]
  print(f"status :{status}")
  if status == True:
    bucket = storage.bucket("shiftmaker-fee68.appspot.com")
    blob = bucket.blob(f"Groups/{groupId}/{file_name}")
    file_data = blob.download_as_string()

    wb = load_workbook(filename=BytesIO(file_data))
    ws = wb.active
    
    val1 = True
    count1= 3
    val2=True
    count2=4

    while val2 ==True:
      if ws.cell(row=count2,column=1).value !=None:
          if ws.cell(row=count2,column=1).value == username:
            while val1 == True:
              if ws.cell(row=2,column=count1).value !=None:
                day=str(ws.cell(row=2,column=count1).value)
                ws.cell(row=count2,column=count1).value=request["開始"][day]
                ws.cell(row=count2+1,column=count1).value=request["終了"][day]
              else:
                val1 = False
              count1+=1
      else:
        val2 = False
      count2+=2

    output_buffer = BytesIO()
    wb.save(output_buffer)
    modified_file_data = output_buffer.getvalue()

    # 修正後のExcelファイルを上書き保存する
    new_blob = bucket.blob(f"Groups/{groupId}/{file_name}")
    new_blob.upload_from_string(modified_file_data)
    membersData=db.collection("Groups").document(groupId).collection("groupInfo").document("member").get().to_dict()

    for key, value in membersData.items():
      print(key)
      print(value["uid"])
      if value["uid"]== userId:
         print("ok")
         db.collection("Groups").document(groupId).collection("groupInfo").document("member").update(
            {f"{key}":{
              "name":f"{value["name"]}",
              "situation":"done",
              "uid":f"{value["uid"]}"}
            })
         break       
  else:
    print("募集期間外")




    
@on_document_written(document="Groups/{groupId}/groupInfo/status", region="asia-northeast1")
def download_link(
    event: Event[Change[DocumentSnapshot]],
  ) -> None:
    Data = event.data.after.to_dict()
    status = Data["status"]
    groupId = event.params["groupId"]
    groupName = db.collection("Groups").document(groupId).collection("groupInfo").document("pass").get().to_dict()["groupName"]
    tokens = []
    member = db.collection("Groups").document(groupId).collection("groupInfo").document("member").get().to_dict()
    for i in range(1,len(member)+1,1):
      uid =member[f"{i}"]["uid"]
      token = db.collection("Users").document(uid).collection("MyInfo").document("userInfo").get().to_dict()["token"]
      print(token)
      tokens.append(token)
      print(f"{tokens}")
    if status==False:
      print(f"false {status}")
      
      
      bucket = storage.bucket("shiftmaker-fee68.appspot.com")
      file_name = db.collection("Groups").document(groupId).collection("groupInfo").document("fileName").get().to_dict()["file_name"]
      print(file_name)
      print(groupId)
      
      blob = bucket.blob(f"Groups/{groupId}/{file_name}")
      url = blob.generate_signed_url(
        method="GET",expiration=datetime.timedelta(days=15))
      print(f"This is a downloadlink:{url}")
      doc_ref = db.collection("Groups").document(groupId).collection('groupInfo').document("download")
      doc_ref.update({'downloadLink': url})
      
      
        
      message = firebase_admin.messaging.MulticastMessage(
         
          tokens=tokens,
          android=firebase_admin.messaging.AndroidConfig(
            notification=firebase_admin.messaging.AndroidNotification(
              title=f'{groupName} からアナウンス',
              body='シフトの募集が締め切られました',
            ),
            priority="high",
            direct_boot_ok=True
          ),
      )
    else:
      print(f"true {status}")
      
      message = firebase_admin.messaging.MulticastMessage(
          notification=firebase_admin.messaging.Notification(
            title=f'{groupName} からアナウンス',
            body='シフト募集開始'
          ),
          tokens=tokens,
          android=firebase_admin.messaging.AndroidConfig(
             priority="high",
             direct_boot_ok=True
          ),
      )
      
    response = firebase_admin.messaging.send_each_for_multicast(message)
    print('Successfully sent message:', response)
    
    

@https_fn.on_call(region="asia-northeast1")
def send_shift(req: https_fn.Request) -> https_fn.Response:
    userId = req.data['userId']
    groupId = req.data['groupId']
    groupName = req.data["groupName"]

    adminId = db.collection("Groups").document(groupId).collection("groupInfo").document("admin").get().to_dict()["adminId"]
    if adminId == userId:      
      shiftData = db.collection("Groups").document(groupId).collection("groupInfo").document("RequestShiftList").get().to_dict()
      for i in shiftData.keys():
        start = shiftData[i]["start"]
        end = shiftData[i]["end"]
        newItem = {}
        for j in start.keys():
          text = f"{start[j]} - {end[j]}"
          dateStr = j.replace("-","/")
          newItem[dateStr] = text
        destinationRef = db.collection("Users").document(i).collection("CompletedShift").document("shift")
        prevData=destinationRef.get().to_dict()[groupName]
        completedShift = {**newItem, **prevData}
        destinationRef.update({groupName:completedShift})




    

#以下はテスト用関数   
"""
@on_document_written(document="makedir/{testId}",
                     region="asia-northeast1")
def test_make_directory(
    event: Event[Change[DocumentSnapshot]],
  ) -> None:
    print("request is arrived")
    Data = event.data.after.to_dict()
    
    userId = Data['userId']
    fullname = Data['fullname']
    time.sleep(2)
    print("request is arrived")
    print(userId)
    print(fullname)
    db.collection("Users").document(userId).collection("MyInfo").document("userInfo").set({"username": fullname,"groupId":{"1":"no data"},"birthday":"2000-01-01","hourlyWage":{"no data":1000},"tokens":"no data"})
    db.collection("Users").document(userId).collection("MyInfo").document("currentGroupNum").set({"currentNum":"1"})
    db.collection("Users").document(userId).collection("MyInfo").document("groups").set({"1":{"groupName":"no data","groupId":"no data"}})
    db.collection("Users").document(userId).collection("RequestShift").document("shift").set({"no data":{"開始": "no data", "終了": "no data"}})
    db.collection("Users").document(userId).collection("CompletedShift").document("shift").set({"no data":{"2000/01/01":""}})


@on_document_written(document="creategroup/{testId}",
                     region="asia-northeast1")
def test_create_group(
    event: Event[Change[DocumentSnapshot]],
  ) -> None:
    Data = event.data.after.to_dict()
    print("request is arrived")
    
    userId = Data["userId"]
    groupName =Data["groupName"]
    groupPass =Data["pass"]
    groupId =uuid.uuid4()
    print(type(groupId))
    print(type(str(groupId)))
    groupId = str(groupId)
    print(f"userId:{userId} groupName:{groupName} grouPass:{groupPass} groupId:{groupId}")

    adminname=db.collection("Users").document(userId).collection("MyInfo").document("userInfo").get().to_dict()["username"]
    group_ref =db.collection("Users").document(userId).collection("MyInfo").document("groups")
    group_data =group_ref.get().to_dict()
    firstGroup = group_data["1"]

    print(f"group_data:{group_data}")
    print(f"group_data[groupId]:{firstGroup["groupId"]}")
    print(firstGroup["groupId"]=="no data")
    

    if firstGroup["groupId"]=="no data":
      groupCount="1"
    else:
      groupCount=str(len(group_data)+1)
      print(groupCount)
    group_ref.update({groupCount:{"groupName":groupName,"groupPass":groupPass,"groupId":groupId}})
      
    doc_ref_groupInfo = db.collection('Groups').document(groupId).collection("groupInfo")

    doc_ref_groupInfo.document("pass").set({"pass": groupPass,"groupName":groupName,"adminName":adminname})
    doc_ref_groupInfo.document("download").set({"downloadLink": "no data"})
    doc_ref_groupInfo.document("status").set({"status":False})
    doc_ref_groupInfo.document("admin").set({"adminId":userId})
    doc_ref_groupInfo.document("fileName").set({"file_name":"no data","comFileName":"no data"})
    doc_ref_groupInfo.document("member").set({"1": "no data"})
    doc_ref_groupInfo.document("tableRequest").set({"start":"no data","end":"no data"})
    doc_ref_groupInfo.document("applicants").set({"1":"no data"})
    doc_ref_groupInfo.document("message").set({"message":""})


@on_document_written(document="deletegroup/{testId}",
                     region="asia-northeast1")
def test_delete_group(
    event: Event[Change[DocumentSnapshot]],
  ) -> None:
  Data = event.data.after.to_dict()
  userId = Data["userId"]
  groupId =Data["groupId"]
  selectedNumber =int(Data["num"])
  print(f"userId: {userId}, groupId: {groupId},selectedNumber:{selectedNumber}")
  doc_ref_delete = db.collection("Groups").document(groupId).collection("groupInfo")
  MyInfo_ref=db.collection("Users").document(userId).collection("MyInfo")
  groups=MyInfo_ref.document("groups").get().to_dict()
  member = doc_ref_delete.document("member").get().to_dict()
  print(f"member:{member}")
  member_count= len(member)
  print(f"member_count:{member_count}")


  new_data = {} 
  for num, data in groups.items():
    if int(num) != selectedNumber:
        print(f" {num}  {data}")
        if selectedNumber ==1:
            num = str(int(num) - 1)
            new_data[num]=data
        elif selectedNumber ==2:
            if num =="3":
                num = str(int(num) - 1)
            new_data[num]=data
        elif selectedNumber ==3:
            new_data[num]=data
  print(f"new_data:{new_data}")

  for i in range(1,member_count+1,1):
    memberUid = member[f"{i}"]["uid"]

    print(memberUid)
    userInfo_ref = db.collection("Users").document(memberUid).collection("MyInfo").document("userInfo")
    userGroupIdData = userInfo_ref.get().to_dict()["groupId"]
    groupIdCount = len(userGroupIdData)
    
    newUserGroupIdData = {}
    borderSwitch = False
    print(groupId in userGroupIdData.values())
    if groupId in userGroupIdData.values():
      for i in range(1,groupIdCount+1,1):
         print(userGroupIdData[f"{i}"])
         if userGroupIdData[f"{i}"] != groupId:
            if borderSwitch == False:
              newUserGroupIdData[f"{i}"] =userGroupIdData[f"{i}"]
            if borderSwitch == True:
              newUserGroupIdData[f"{i-1}"] =userGroupIdData[f"{i}"]               
         else:
            borderSwitch = True
            print("ooo")
      if newUserGroupIdData.get("1") is None:
        print("bbb")
        newUserGroupIdData["1"] = "no data"
      userInfo_ref.update({"groupId":newUserGroupIdData})
    print("done")

  
  if len(new_data)==0:
      new_data["1"]={"groupId":"no data","groupName":"no data"}
  MyInfo_ref.document("currentGroupNum").update({"currentNum":"1"})
  MyInfo_ref.document("groups").set(new_data)
  if groupId !="no data":
    doc_ref_delete.document("admin").delete()
    doc_ref_delete.document("applicants").delete()
    doc_ref_delete.document("download").delete()
    doc_ref_delete.document("fileName").delete()
    doc_ref_delete.document("member").delete()
    doc_ref_delete.document("message").delete()
    doc_ref_delete.document("pass").delete()
    doc_ref_delete.document("status").delete()
    doc_ref_delete.document("tableRequest").delete()
    db.collection("Groups").document(groupId).delete()



@on_document_written(document="requestgroup/{testId}",
                     region="asia-northeast1")
def test_request_group(
    event: Event[Change[DocumentSnapshot]],
  ) -> None:
    Data = event.data.after.to_dict()

    userId = Data['userId']
    groupId = Data['groupId']

    userInfo_ref = db.collection("Users").document(userId).collection("MyInfo").document("userInfo")
    username =userInfo_ref.get().to_dict()["username"]
    
    applicants_ref = db.collection("Groups").document(groupId).collection("groupInfo").document("applicants")
    applicants_data=applicants_ref.get().to_dict()
    if applicants_data["1"] == "no data":
      applicants_count=str(len(applicants_data))
      print("no data")
      applicants_ref.update({applicants_count:{"name":username, "situation":"not yet", "uid":userId}})
    else:   
      applicants_count=str(len(applicants_data)+1)
      print(f"already :{applicants_count}")
      applicants_ref.update({applicants_count:{"name":username, "situation":"not yet", "uid":userId}})

#テスト用関数
@on_document_written(document="writemember/{testId}",
                     region="asia-northeast1")
def test_writeMemberInfo(
    event: Event[Change[DocumentSnapshot]],
  ) -> None:
    Data = event.data.after.to_dict()

    userId = Data['userId']
    groupId = Data['groupId']

    userInfo_ref = db.collection("Users").document(userId).collection("MyInfo").document("userInfo")
    username =userInfo_ref.get().to_dict()["username"]
    
    member_ref = db.collection("Groups").document(groupId).collection("groupInfo").document("member")
    member_data=member_ref.get().to_dict()
    if member_data["1"] == "no data":
      member_count=str(len(member_data))
      print("no data")
      member_ref.update({member_count:{"name":username, "situation":"not yet", "uid":userId}})
    else:   
      member_count=str(len(member_data)+1)
      print(f"already :{member_count}")
      member_ref.update({member_count:{"name":username, "situation":"not yet", "uid":userId}})

@on_document_written(document="admitmember/{testId}",
                     region="asia-northeast1")
def test_admit_member(
    event: Event[Change[DocumentSnapshot]],
  ) -> None:
    Data = event.data.after.to_dict()
    userId = Data['userId']
    groupId = Data['groupId']
    groupName = Data['groupName']

    userInfo_ref = db.collection("Users").document(userId).collection("MyInfo").document("userInfo")
    userGroupIdData = userInfo_ref.get().to_dict()["groupId"]
    groupIdCount = len(userGroupIdData)
    userHourlyWage = userInfo_ref.get().to_dict()["hourlyWage"]
     
    print(f"userGroupIdData : \n {userGroupIdData}")
    print(f"groupIdCount : \n {groupIdCount}")
    print(f"hourlyWage : \n {userHourlyWage}")
    if groupIdCount == 1 and userGroupIdData["1"] == "no data":
      userGroupIdData["1"] = groupId
    else:
      userGroupIdData[f"{groupIdCount+1}"] = groupId
    
    switch = True
    i = 1 
    while switch:
      if groupName in userHourlyWage:
        groupName = groupName + f"({i})"
        i = i + 1
      else:
        switch = False
            
    userHourlyWage[f"{groupName}"] = 1000
    userInfo_ref.update({"groupId":userGroupIdData,"hourlyWage":userHourlyWage})


    return https_fn.Response(f"ok")


@on_document_written(document="deletemember/{testId}",
                     region="asia-northeast1")
def test_delete_member(
    event: Event[Change[DocumentSnapshot]],
  ) -> None:
    Data = event.data.after.to_dict()
    userId = Data['userId']
    groupId = Data['groupId']

    userInfo_ref = db.collection("Users").document(userId).collection("MyInfo").document("userInfo")
    userGroupIdData = userInfo_ref.get().to_dict()["groupId"]
    groupIdCount = len(userGroupIdData)

    print(f"userGroupIdData : \n {userGroupIdData}")
    print(f"groupIdCount : \n {groupIdCount}")
    
    newUserGroupIdData = {}
    borderSwitch = False
    print(groupId in userGroupIdData.values())
    if groupId in userGroupIdData.values():
      for i in range(1,groupIdCount+1,1):
         print(userGroupIdData[f"{i}"])
         if userGroupIdData[f"{i}"] != groupId:
            if borderSwitch == False:
              newUserGroupIdData[f"{i}"] =userGroupIdData[f"{i}"]
            if borderSwitch == True:
              newUserGroupIdData[f"{i-1}"] =userGroupIdData[f"{i}"]               
         else:
            borderSwitch = True
            print("ooo")
      if newUserGroupIdData.get("1") is None:
        print("bbb")
        newUserGroupIdData["1"] = "no data"
      userInfo_ref.update({"groupId":newUserGroupIdData})
    print("done")
"""


"""
@on_document_written(document="Users/{userId}/RequestShift/shift")
def input_data(
    event: Event[Change[DocumentSnapshot]],
  ) -> None:


  userId = event.params["userId"]
  docData = event.data.after.to_dict()
  groupId = docData["RequestedGroupId"]
  request = docData[groupId]
  userInfo = db.collection("Users").document(userId).collection("MyInfo").document("userInfo").get().to_dict()
  username = userInfo["username"]
  file_name = db.collection("Groups").document(groupId).collection("groupInfo").document("fileName").get().to_dict()["file_name"]
  print(file_name)
  status = db.collection("Groups").document(groupId).collection('groupInfo').document("status").get().to_dict()["status"]
  print(f"status :{status}")
  if status == True:
    bucket = storage.bucket("shiftmaker-fee68.appspot.com")
    blob = bucket.blob(f"Groups/{groupId}/{file_name}")
    file_data = blob.download_as_string()

    wb = load_workbook(filename=BytesIO(file_data))
    ws = wb.active
    
    val1 = True
    count1= 3
    val2=True
    count2=4

    while val2 ==True:
      if ws.cell(row=count2,column=1).value !=None:
          if ws.cell(row=count2,column=1).value == username:
            while val1 == True:
              if ws.cell(row=2,column=count1).value !=None:
                day=str(ws.cell(row=2,column=count1).value)
                ws.cell(row=count2,column=count1).value=request["開始"][day]
                ws.cell(row=count2+1,column=count1).value=request["終了"][day]
              else:
                val1 = False
              count1+=1
      else:
        val2 = False
      count2+=2

    output_buffer = BytesIO()
    wb.save(output_buffer)
    modified_file_data = output_buffer.getvalue()

    # 修正後のExcelファイルを上書き保存する
    new_blob = bucket.blob(f"Groups/{groupId}/{file_name}")
    new_blob.upload_from_string(modified_file_data)
    membersData=db.collection("Groups").document(groupId).collection("groupInfo").document("member").get().to_dict()

    for key, value in membersData.items():
      print(key)
      print(value["uid"])
      if value["uid"]== userId:
         print("ok")
         db.collection("Groups").document(groupId).collection("groupInfo").document("member").update(
            {f"{key}":{
              "name":f"{value["name"]}",
              "situation":"done",
              "uid":f"{value["uid"]}"}
            })
         break       
  else:
    print("募集期間外")


"""
"""
    
@storage_fn.on_object_finalized(bucket="shiftmaker-fee68.appspot.com",
                               region="asia-northeast1")
def scan_excel_file(event: storage_fn.CloudEvent[storage_fn.StorageObjectData]):

  name = event.data.name
  groupId = name.split('/')[1]
  if name.split('/')[2] == "completed":
    
    bucket = storage.bucket("shiftmaker-fee68.appspot.com")

    blob = bucket.blob(name)
    new_file_name = name.split("/")[-1]
    
    member=db.collection("Groups").document(groupId).collection("groupInfo").document("member").get().to_dict()
    groupName =db.collection("Groups").document(groupId).collection("groupInfo").document("pass").get().to_dict()["groupName"]

   #ユーザーへの通知メッセージの作成
    tokens = []
    for i in range(1,len(member)+1,1):
      uid =member[f"{i}"]["uid"]
      token = db.collection("Users").document(uid).collection("MyInfo").document("userInfo").get().to_dict()["token"]
      print(token)
      tokens.append(token)
      print(f"{tokens}")
    message = firebase_admin.messaging.MulticastMessage(
          notification=firebase_admin.messaging.Notification(
            title=f'{groupName} からアナウンス',
            body='新規シフトが確定しました'
          ),
          tokens=tokens,
          android=firebase_admin.messaging.AndroidConfig(
             priority="high",
             direct_boot_ok=True
          ),
      )
    
    file_data = blob.download_as_string()
    wb = load_workbook(filename=BytesIO(file_data))
    ws = wb.active

    count1= 3
    count2=4
    val= True
    shift_list=[]
    firstDay = datetime.datetime.strptime(ws.cell(row=1,column=1).value,"%Y/%m/%d")
    lastDay = datetime.datetime.strptime(ws.cell(row=1,column=2).value,"%Y/%m/%d")
    duration= (lastDay-firstDay).days
    while val:
      shiftData ={}
      if ws.cell(row=count2,column=1).value !=None:
          data = {"username":"","shift":shiftData,"uid":""}
          data["username"]= ws.cell(row=count2,column=1).value
          for i in range(1,len(member)+1,1):
            if data["username"]==member[f"{i}"]["name"]:
              data["uid"] = member[f"{i}"]["uid"]
              break
          for i in range(0,duration+1,1):
            start=ws.cell(row=count2,column=count1+i).value
            end = ws.cell(row=count2+1,column=count1+i).value

            if start != "-" or end != "-":
              start=str(start).split(":")
              end = str(end).split(":")
              start=f"{start[0]}:{start[1]}"
              end=f"{end[0]}:{end[1]}"
              shiftData[firstDay.strftime("%Y/%m/%d")] = f"{start} - {end}"
            firstDay= firstDay+datetime.timedelta(days=1)
          shift_list.append(data)
          print(data)
      else:
        val=False
      count2+=2
    
    for data in shift_list:
      if not data["shift"]:
        print(f"{data["uid"]}のシフトが空です")
      else:
        newShift = db.collection("Users").document(data["uid"]).collection("CompletedShift").document("shift").get().to_dict()[f"{groupName}"]
        newShift.update(data["shift"])
        print(newShift)
        db.collection("Users").document(data["uid"]).collection("CompletedShift").document("shift").update({f"{groupName}":newShift})

    db.collection("Groups").document(groupId).collection("groupInfo").document("fileName").update({"comFileName":new_file_name})
    
    response = firebase_admin.messaging.send_each_for_multicast(message)
    print('Successfully sent message:', response)

"""