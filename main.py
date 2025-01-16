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
  on_document_written,gi
  Event,
  DocumentSnapshot,
  Change 
)


cred = credentials.Certificate('')
firebase_admin.initialize_app(cred, {
    'storageBucket': ''
})
db = firestore.client()

#今後の課題
#ユーザー名、グループ名が被った場合の処理が不完全
#現状だと被った場合に処理が止まる可能性が高い
#応急処置としてadmit_memberで時給リストに保存するグループ名は被らなくしている


#ユーザー登録時にFirestoreにコレクション/ドキュメントを製作する関数
@https_fn.on_call(region="asia-northeast1")
def make_directory(req: https_fn.CallableRequest):
    #ユーザーIDの取得
    userId = req.data['userId']
    #ユーザー名の取得
    fullname = req.data['fullname']
    #取得した情報を表示
    print("request is arrived")
    print(userId)
    print(fullname)
    #コレクション/ドキュメントの製作
    db.collection("Users").document(userId).collection("MyInfo").document("userInfo").set({"username": fullname,"groupId":{"1":"no data"},"birthday":"2000-01-01","hourlyWage":{"no data":1000},"tokens":"no data"})
    db.collection("Users").document(userId).collection("MyInfo").document("currentGroupNum").set({"currentNum":"1"})
    db.collection("Users").document(userId).collection("MyInfo").document("groups").set({"1":{"groupName":"no data","groupId":"no data"}})
    db.collection("Users").document(userId).collection("RequestShift").document("shift").set({"no data":{"開始": "no data", "終了": "no data"}})
    db.collection("Users").document(userId).collection("CompletedShift").document("shift").set({"no data":{"2000/01/01":"12:00 - 12:00"}})
    

#新たなグループの製作を行う関数
@https_fn.on_call(region="asia-northeast1")
def create_group(req: https_fn.CallableRequest):
    #ユーザーIDの取得    
    userId = req.data["userId"]
    #グループ名の取得
    groupName =req.data["groupName"]
    #グループパスワードの取得
    groupPass =req.data["pass"]
    #グループIDの生成
    groupId =uuid.uuid4()
    #文字列に変換
    groupId = str(groupId)
    #登録に使う情報を一覧表示
    print(f"userId:{userId} groupName:{groupName} grouPass:{groupPass} groupId:{groupId}")
    #申請者の名前（グループの管理者になる）
    adminname=db.collection("Users").document(userId).collection("MyInfo").document("userInfo").get().to_dict()["username"]
    #管理中のグループを一覧表示
    group_ref =db.collection("Users").document(userId).collection("MyInfo").document("groups")
    group_data =group_ref.get().to_dict()
    print(f"group_data:{group_data}")
    #管理中グループの先頭を表示
    #管理中のグループの先頭がno dataの場合は管理グループは1つ（初期状態）
    #no data が入っているため、そのままやるとグループカウントが２になってしまう
    #これを回避するためにno dataの場合にグループカウントが1になるようにする
    firstGroup = group_data["1"]
    print(f"group_data[groupId]:{firstGroup["groupId"]}")
    #if文の分岐を表示
    print(firstGroup["groupId"]=="no data")
    if firstGroup["groupId"]=="no data":
      groupCount="1"
    else:
      groupCount=str(len(group_data)+1)
      print(groupCount)
    #管理者のドキュメントにあるグループの情報に新規グループを加えて更新する
    group_ref.update({groupCount:{"groupName":groupName,"groupPass":groupPass,"groupId":groupId}})
    #新規グループの参照を取得
    doc_ref_groupInfo = db.collection('Groups').document(groupId).collection("groupInfo")
    #新規グループのドキュメントを製作
    doc_ref_groupInfo.document("pass").set({"pass": groupPass,"groupName":groupName,"adminName":adminname})
    doc_ref_groupInfo.document("download").set({"downloadLink": "no data"})
    doc_ref_groupInfo.document("status").set({"status":False})
    doc_ref_groupInfo.document("admin").set({"adminId":userId})
    doc_ref_groupInfo.document("fileName").set({"file_name":"no data","comFileName":"no data"})
    doc_ref_groupInfo.document("member").set({"1": "no data"})
    doc_ref_groupInfo.document("tableRequest").set({"start":"no data","end":"no data"})
    doc_ref_groupInfo.document("applicants").set({"1":"no data"})
    doc_ref_groupInfo.document("message").set({"message":""})

#管理グループを削除する関数
@https_fn.on_call(region="asia-northeast1")
def delete_group(req: https_fn.CallableRequest):
  #ユーザーIDの取得
  userId = req.data["userId"]
  #グループIDの取得
  groupId =req.data["groupId"]
  #削除するグループのインデックスを取得
  selectedNumber =int(req.data["num"])
  #取得情報を一覧表示
  print(f"userId: {userId}, groupId: {groupId},selectedNumber:{selectedNumber}")
  #削除するグループの参照取得
  doc_ref_delete = db.collection("Groups").document(groupId).collection("groupInfo")
  #管理者のコレクションへの参照取得
  MyInfo_ref=db.collection("Users").document(userId).collection("MyInfo")
  #管理中グループの一覧を保存しているドキュメントのデータを取得
  groups=MyInfo_ref.document("groups").get().to_dict()
  #削除後の更新に取得するための箱
  new_data = {} 
  #現状管理できるグループ数は3つまでにしている
  #削除するグループのインデックスによって分岐
  #削除するグループ以外の要素だけを残す
  for num, data in groups.items():
    if int(num) != selectedNumber:
        print(f" {num}  {data}")
        #削除するインデックスが1
        if selectedNumber ==1:
            num = str(int(num) - 1)
            new_data[num]=data
        #削除するインデックスが2
        elif selectedNumber ==2:
            if num =="3":
                num = str(int(num) - 1)
            new_data[num]=data
        #削除するインデックスが3
        elif selectedNumber ==3:
            new_data[num]=data
  #更新するデータの状態によって分岐
  #全てのグループが削除された場合はno dataを入れる
  if len(new_data)==0:
      new_data["1"]={"groupId":"no data","groupName":"no data"}
  #選択中のグループのインデックスを１に再設定
  MyInfo_ref.document("currentGroupNum").update({"currentNum":"1"})
  #更新データをセット
  MyInfo_ref.document("groups").set(new_data)
  #削除するグループがno dataではないか念のため検証
  #実際のグループのコレクション/ドキュメントを削除
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

#グループへの参加申請を行う関数
@https_fn.on_call(region="asia-northeast1")
def request_group(req: https_fn.CallableRequest):
    #ユーザーIDを取得
    userId = req.data['userId']
    #グループIDを取得
    groupId = req.data['groupId']
    #取得情報を表示
    print(f"userId: {userId}, groupId: {groupId}")
    #リクエスト側のユーザー情報への参照を取得
    userInfo_ref = db.collection("Users").document(userId).collection("MyInfo").document("userInfo")
    #参照からユーザー名を取得
    username =userInfo_ref.get().to_dict()["username"]
    #グループの申請者リストを保存しているドキュメントへの参照を取得
    applicants_ref = db.collection("Groups").document(groupId).collection("groupInfo").document("applicants")
    #申請者データ取得
    applicants_data=applicants_ref.get().to_dict()
    #申請者がいるかいないかで分岐
    #先頭がno dataの場合は申請者なしだからカウントは据え置き
    #申請者がいる場合はカウントを増やす
    #申請者データを更新
    if applicants_data["1"] == "no data":
      applicants_count=str(len(applicants_data))
      print("no data")
      applicants_ref.update({applicants_count:{"name":username, "situation":"not yet", "uid":userId}})
    else:   
      applicants_count=str(len(applicants_data)+1)
      print(f"already :{applicants_count}")
      applicants_ref.update({applicants_count:{"name":username, "situation":"not yet", "uid":userId}})


#申請者を承認するための関数
@https_fn.on_call(region="asia-northeast1")
def admit_member(req: https_fn.Request) -> https_fn.Response:
  #ユーザーIDを取得
  userId = req.data['userId']
  #グループIDを取得
  groupId = req.data['groupId']
  #グループ名を取得
  groupName = req.data['groupName']
  #取得情報を一覧表示
  print(f"userId: {userId}, groupId: {groupId},groupName:{groupName}")
  #追加するメンバーのユーザー情報への参照を取得
  userInfo_ref = db.collection("Users").document(userId).collection("MyInfo").document("userInfo")
  #追加するメンバーの確定シフトへの参照を取得
  completedShift_ref = db.collection("Users").document(userId).collection("CompletedShift").document("shift")
  #追加するメンバーのグループIDリストの取得
  userGroupIdData = userInfo_ref.get().to_dict()["groupId"]
  #グループIDの数
  groupIdCount = len(userGroupIdData)
  #追加するメンバーの時給リストを取得
  userHourlyWage = userInfo_ref.get().to_dict()["hourlyWage"]
  #ユーザーのグループIDリストを更新するための処理
  if groupIdCount == 1 and userGroupIdData["1"] == "no data":
    #追加するメンバーがどのグループにも所属していない場合
    userGroupIdData["1"] = groupId
  else:
    #既に他のグループに所属している場合
    userGroupIdData[f"{groupIdCount+1}"] = groupId
  
  #グループ名が被った場合の処理(応急処置)
  switch = True
  i = 1 
  while switch:
    if groupName in userHourlyWage:
      #グループ名の後ろに数字を付ける
      groupName = groupName + f"({i})"
      i = i + 1
    else:
      switch = False
  #初期時給を1000円とする
  userHourlyWage[f"{groupName}"] = 1000
  #確定シフトを格納するための場所を用意する
  completedShift_ref.update({f"{groupName}":{"2000/01/01":"12:00 - 12:00"}})
  #グループIDリストと時給リストを更新
  userInfo_ref.update({"groupId":userGroupIdData,"hourlyWage":userHourlyWage})


#グループ所属メンバーを削除する際の関数
#グループのメンバーリストを削除するのではなく、削除対象のグループIDリストを編集する関数
@https_fn.on_call(region="asia-northeast1")
def delete_member(req: https_fn.Request):
  #ユーザーIDの取得
  userId = req.data['userId']
  #グループIDの取得
  groupId = req.data['groupId']
  #削除ユーザーのユーザー情報への参照を取得
  userInfo_ref = db.collection("Users").document(userId).collection("MyInfo").document("userInfo")
  #グループIDリストを取得
  userGroupIdData = userInfo_ref.get().to_dict()["groupId"]
  #所属しているグループの数を取得
  groupIdCount = len(userGroupIdData)
  #取得情報を一覧表示
  print(f"userId : {userId} groupId : {groupId} userGroupIdData : {userGroupIdData} groupIdCount : {groupIdCount}")
  #更新するためのグループIDを入れる箱
  newUserGroupIdData = {}
  #該当グループIDを通過したかの切換を行うための真偽値
  borderSwitch = False
  #グループIDリストの中に該当グループIDがあるかどうかを検証
  if groupId in userGroupIdData.values():
    #グループIDを1つずつ取り出す
    for i in range(1,groupIdCount+1,1):
        print(userGroupIdData[f"{i}"])
        #取り出したグループIDが該当グループと一致するか検証
        if userGroupIdData[f"{i}"] != groupId:
          #一致しなかった場合
          if borderSwitch == False:
            #該当グループIDを通過していない場合は更新データにそのまま代入
            newUserGroupIdData[f"{i}"] =userGroupIdData[f"{i}"]
          if borderSwitch == True:
            #該当グループIDを通過した後の場合は更新データのインデックスを繰り上げて代入
            #キーとしているインデックスに隙間が出来ないようにするため
            newUserGroupIdData[f"{i-1}"] =userGroupIdData[f"{i}"]               
        else:
          #一致した場合
          #真偽値を切換
          borderSwitch = True
    #所属しているグループがなくなった場合
    #グループIDの先頭にno dataを代入
    if newUserGroupIdData.get("1") is None:
      newUserGroupIdData["1"] = "no data"
    #グループIDリストを更新
    userInfo_ref.update({"groupId":newUserGroupIdData})
  #処理の終了を表示
  print("done")


#ユーザーに 募集/停止 通知を送るための関数
@on_document_written(document="Groups/{groupId}/groupInfo/status", region="asia-northeast1")
def status_notification(
    event: Event[Change[DocumentSnapshot]],
  ) -> None:
    #statusドキュメントの内容を取得
    Data = event.data.after.to_dict()
    #statusを取得(真偽値)
    status = Data["status"]
    #グループIDの取得
    groupId = event.params["groupId"]
    #グループ名リストの取得
    groupName = db.collection("Groups").document(groupId).collection("groupInfo").document("pass").get().to_dict()["groupName"]
    #FCMトークンを入れる箱
    tokens = []
    #グループのメンバーリストを取得
    member = db.collection("Groups").document(groupId).collection("groupInfo").document("member").get().to_dict()
    #メンバーリストからFCMトークンを一つづつ取り出し、リストに追加
    for i in range(1,len(member)+1,1):
      uid =member[f"{i}"]["uid"]
      token = db.collection("Users").document(uid).collection("MyInfo").document("userInfo").get().to_dict()["token"]
      print(token)
      tokens.append(token)
      print(f"{tokens}")
    #ステータスの状態によって分岐
    if status==False:
      #シフト募集が停止中の場合
      #FCMトークンのリストに含まれる端末に一斉に通知を送信するメソッド
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
      #シフト募集が募集中の場合
      #FCMトークンのリストに含まれる端末に一斉に通知を送信するメソッド
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
    #メッセージ送信
    response = firebase_admin.messaging.send_each_for_multicast(message)
    print('Successfully sent message:', response)
    
    

#確定シフトデータをグループメンバーに配る関数
@https_fn.on_call(region="asia-northeast1")
def send_shift(req: https_fn.Request):
    #ユーザーIDの取得
    userId = req.data['userId']
    #グループIDの取得
    groupId = req.data['groupId']
    #グループ名の取得
    groupName = req.data["groupName"]
    #グループIDから管理者のIDを取得する
    adminId = db.collection("Groups").document(groupId).collection("groupInfo").document("admin").get().to_dict()["adminId"]
    #グループIDと操作を行ったユーザーのIDが同じか検証
    if adminId == userId:
      #グループに保存してある確定シフトを取得
      shiftData = db.collection("Groups").document(groupId).collection("groupInfo").document("RequestShiftList").get().to_dict()
      #シフトデータのキーはユーザーIDになっている
      #ユーザー１人ずつのシフトデータを成形してからユーザーごと
      #のドキュメントに上書き
      for i in shiftData.keys():
        #シフトデータの取り出し
        start = shiftData[i]["start"]
        end = shiftData[i]["end"]
        #更新データを入れる箱
        newItem = {}
        #シフトデータのキーは日付（文字列）になっている
        for j in start.keys():
          #シフトの時間を取り出し合わせる
          text = f"{start[j]} - {end[j]}"
          #日付の書き方を変更  例：(2025-01-01 -> 2025/01/01)
          dateStr = j.replace("-","/")
          #シフトデータを格納
          newItem[dateStr] = text
        #ユーザーごとの確定シフトへの参照を取得
        destinationRef = db.collection("Users").document(i).collection("CompletedShift").document("shift")
        #該当グループのシフトを取得
        #グループ名は衝突する可能性があるため、修正を検討中
        prevData=destinationRef.get().to_dict()[groupName]
        #新規シフトに旧シフトを結合する
        completedShift = {**newItem, **prevData}
        #確定シフトを更新
        destinationRef.update({groupName:completedShift})

