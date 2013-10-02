'    rename.vbs
'    ��Ւn�}�����{��t�@�C����->�A���t�@�x�b�g�t�@�C�����ϊ�
'
'---------------------
'    begin                : September 2013
'    copyright            : (C) 2013 by Yoichi Kayama
'    email                : yoichi.kayama at gmail dot com
'***************************************************************************
'*                                                                         *
'*   This program is free software; you can redistribute it and/or modify  *
'*   it under the terms of the GNU General Public License as published by  *
'*   the Free Software Foundation; either version 2 of the License, or     *
'*   (at your option) any later version.                                   *
'*                                                                         *
'***************************************************************************/
'
Option Explicit
On Error Resume Next


Dim objApl
Dim objFolder
Dim objFolderItems
Dim objItem
Dim  i
dim  objFS
Dim  scStr
dim  newName
dim ic

Dim replaced
Dim  jfNameAr, asNameAr
Dim result

jfNameAr = _
Array("�s�����E��","�s������\�_","������",_
"�����E��","�����̑�\�_","����","���U��","�����\������", _
"�����\������","���z���̊O����","���H��", _
"���H�\����","�W���_","�O���̒��S��" ,"�s�����","���z��","�C�ݐ�")

asNameAr = _
Array("AdmBdy","AdmPt","Cntr",_
"CommBdy","CommPt","WA","WL","WStrL",_
"WStrA","BldL","RdEdg",_
"RdCompt","ElevPt","RailCL","AdmArea","BldA","Cstline")



Set objApl = CreateObject("Shell.Application")

If Err.Number = 0 Then
   

    Set objFolder = objApl.BrowseForFolder(0, "�ϊ��t�@�C���i�[�t�H���_��I�����ĉ�����", 0, "C:\")
    If Not objFolder Is Nothing Then
       ' WScript.Echo objFolder.Items.Item.Path
    End If

Else
    WScript.Echo "�G���[�F" & Err.Description
End If



Set objFS = CreateObject("Scripting.FileSystemObject")


Set objFolderItems = objFolder.Items()

result = "�@�@�ϊ���ƌ���  "

For i=0 To objFolderItems.Count-1

 
    Set objItem = objFolderItems.Item(i)

    '�@file or folder ����
    If objItem.IsFolder = True Then
     
    Else
       ' file

   

     replaced = 0

     for ic = 0 to 16

      scStr = jfNameAr(ic)

      if InStr( objItem.Name, scStr) > 0 then
        'WScript.Echo "  " & objItem.Path
   
        if replaced = 0 then

             newName = Replace(objItem.Path, scStr, asNameAr(ic))

            
       ' WScript.Echo "  " & objItem.Path & "  " & newName
            Call objFS.CopyFile(objItem.Path, newName)
            result = result &  vbCrLf  & objItem.Path & "->" & newName
            replaced = 1
         end if
      end if

      Next

    End If

  

Next


MsgBox result

set scStr = Nothing
set objFS = Nothing
Set objItem = Nothing
Set objFolderItems = Nothing
Set objFolder = Nothing
Set objApl = Nothing