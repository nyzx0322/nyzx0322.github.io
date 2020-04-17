<%@language="VBScript" CodePage="65001"%>
<%Option Explicit%>
<%Session.CodePage=65001%>
<%Response.Charset="UTF-8"%>
<%Server.ScriptTimeout=60%>
<%dim WebpageSort:WebpageSort="User"%>
<%dim WebpageName:WebpageName=""%>
<%dim WebpageTitle:WebpageTitle=""%>

<%
on error resume next


dim WebsiteName,WebsiteDomain
dim GB_UserName,GB_UserEMail,GB_UserMobilePhone,GB_UserIMQQ,GB_UserIMWeixin,GB_UserIMYixin,GB_UserIMMomo,GB_Content,GB_Reply,GB_Status,GB_Type
dim GuestbookAddDistanceCheck,GuestbookAddDistanceSecond,GuestbookContentLength,GuestbookAddCaptchaShow
dim GuestbookAddEMailEnabled,GuestbookSMTPSubject,GuestbookSMTPAddRecipient,GuestbookSMTPServer,GuestbookSMTPUsername,GuestbookSMTPPassword,GuestbookSMTPBody
dim GuestbookListPageSize,GuestbookIDNext
dim AM_Account,AM_Password,AdministratorModifyCheck,AdministratorPasswordConfirm,AdministratorPasswordModify,AdministratorLogListPageSize,AdministratorModifyBoolean
dim AdministratorID,AdministratorAccount,AdministratorPurview,AdministratorLoginLdentifying,AdministratorLoginCookiesEnabled,AdministratorLoginTimeout,AdministratorLoginCaptchaShow
dim AdministratorLoginRestrictedEnabled,AdministratorLoginRestrictedDistance,AdministratorLoginRestrictedMistake,AdministratorLoginRestrictedMinute,AdministratorLoginRestrictedMessage
dim ACMA,GuDatabaseFileName,GuDatabaseTablePrefix,WebpagePath
dim AA,BB,CC,DD,EE,FF,GG,HH,II,JJ,KK,LL,MM,NN,OO,PP,QQ,RR,SS,TT,UU,VV,WW,XX,YY,ZZ
dim ARAC,ARAS,ARMC,ARMS,ARLR,ARMR,ARRS,ARSA,ARSB,ARAR,ARSW,ARAP,ARSD,ARML,ARMM,ARIP,SQL,SQLKeyword,SQLBofEof,SearchKeywordString,SQLStatus
dim ASFC,ASFR,SFSO,SFFC,SFFG,SFFL,SFFR,SFCF,SCTF
dim GuBrowserIP,GuCaptchaForm,GuResourceSortA,GuResourceSortB,GuResourceID,GuResourceIcon,GuResourceName,GuResourceNameStyle,GuResourceSite,GuRequestName,GuRequestValue
dim PageCount,PageNumber,PageOrder,PagePresent,PageSize,PageStart,PageEnd,PageTotal
dim GuSystemMessageContent,GuBrowserIPAddressKuaidial


WebsiteName="啊估留言簿"  '网页标题

GuestbookAddDistanceSecond=60  '用户留言间隔，单位：秒，默认60秒。等号后面没有双引号，必须是纯数字并且≥1。
GuestbookContentLength=1000  '留言内容最大字符，默认1000个。等号后面没有双引号，必须是纯数字并且≥1。
GuestbookAddCaptchaShow="1"  '用户留言是否需要验证码：0=取消，1=启用
GuestbookAddEMailEnabled="0"  '设置是否启用用户留言自动发送留言到电子邮箱：0=取消，1=启用
GuestbookSMTPSubject="啊估留言簿用户留言提醒邮件"  '发送用户留言的邮件主题
GuestbookSMTPAddRecipient="12345678@qq.com"  '收件人电子邮箱地址
GuestbookSMTPServer="smtp.exmail.qq.com"  '发送邮件的服务器
GuestbookSMTPUsername="12345678@qq.com"  '发送邮件的邮箱账号
GuestbookSMTPPassword="abcdefghi"  '发送邮件的邮箱密码

GuestbookListPageSize=10  '记事留言每页显示的数目。等号后面没有双引号，必须是纯数字并且≥1。

AdministratorLoginCaptchaShow="1"  '登录管理是否需要验证码：0=取消，1=启用
AdministratorLoginCookiesEnabled="1"  '登录会话方式：1=Cookies，2=Session
AdministratorLoginTimeout=60  '登录超时时间，单位：分钟。等号后面没有双引号，必须是纯数字并且≥1。
AdministratorLoginRestrictedEnabled="1"  '是否启用登录限制功能：0=取消，1=启用
AdministratorLoginRestrictedDistance=120  '多少秒钟以内。等号后面没有双引号，必须是纯数字并且≥1。
AdministratorLoginRestrictedMistake=10  '连续登录失败多少次。等号后面没有双引号，必须是纯数字并且≥1。
AdministratorLoginRestrictedMinute=60  '限制IP地址多少分钟。等号后面没有双引号，必须是纯数字并且≥1。
AdministratorLoginRestrictedMessage="登录失败，同一IP地址在2分钟以内，连续登录失败10次，系统自动限制IP地址60分钟。"  '登录限制提示消息。

AdministratorLogListPageSize=15  '管理账号日志每页显示的数目。等号后面没有双引号，必须是纯数字并且≥1。

GuDatabaseFileName="Guestbook.mdb"  '数据库文件路径及名称，正式使用后一定要修改。
GuDatabaseTablePrefix="GUEEON_CN"  '数据库的表名前缀，不清楚的不要修改此项。
GuBrowserIPAddressKuaidial="http://www.ip138.com/ips138.asp?action=2&ip="  'IP地址查询网址。


WebpagePath=trim(Request.ServerVariables("SCRIPT_NAME"))
for FF=len(WebpagePath) to 1 step -1
if mid(WebpagePath,FF,1)="/" then
	WebpageName=right(WebpagePath,len(WebpagePath)-FF)
	exit for
end if
next

if len(trim(Request.ServerVariables("HTTP_X_FORWARDED_FOR")))>=1 then
	GuBrowserIP=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
else
	GuBrowserIP=Request.ServerVariables("REMOTE_ADDR")
end if

if AdministratorLoginCookiesEnabled="1" then
	AdministratorID=Request.Cookies("GUEEONGUESTBOOKADMINISTRATOR")("ID")
	AdministratorAccount=Request.Cookies("GUEEONGUESTBOOKADMINISTRATOR")("ACCOUNT")
	AdministratorLoginLdentifying=Request.Cookies("GUEEONGUESTBOOKADMINISTRATOR")("LOGINLDENTIFYING")
else
	AdministratorID=session("GUEEONGUESTBOOKADMINISTRATORID")
	AdministratorAccount=session("GUEEONGUESTBOOKADMINISTRATORACCOUNT")
	AdministratorLoginLdentifying=session("GUEEONGUESTBOOKADMINISTRATORLOGINLDENTIFYING")
end if
%>

<!DOCTYPE HTML>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html;charset=UTF-8" />
<title><%=WebsiteName%></title>
<meta name="Keywords" content="" />
<meta name="Description" content="" />
<style type="text/css">
<!--
body {font-family:'宋体','新宋体','宋体-方正超大字符集','Arial','Simsun','Times New Roman','Verdana';font-size:16px;}

a{text-transform:none;text-decoration:none;}
a:hover{color:#FF0000;text-decoration:underline;}

.BlackG12 {font-family:"Georgia","Arial","Times New Roman","宋体","新宋体";font-size:12px;color:#000000;}
.BlackG14 {font-family:"Georgia","Arial","Times New Roman","宋体","新宋体";font-size:14px;color:#000000;}
.BlackG16 {font-family:"Georgia","Arial","Times New Roman","宋体","新宋体";font-size:16px;color:#000000;}
.BlackG18 {font-family:"Georgia","Arial","Times New Roman","宋体","新宋体";font-size:18px;color:#000000;}
.BlackS12 {font-family:"宋体","新宋体";font-size:12px;color:#000000;}
.BlackS14 {font-family:"宋体","新宋体";font-size:14px;color:#000000;}
.BlackS16 {font-family:"宋体","新宋体";font-size:16px;color:#000000;}
.BlackS18 {font-family:"宋体","新宋体";font-size:18px;color:#000000;}
.BlackV12 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:12px;color:#000000;}
.BlackV14 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:14px;color:#000000;}
.BlackV16 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:16px;color:#000000;}
.BlackV18 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:18px;color:#000000;}

.BlueG12 {font-family:"Georgia","Arial","Times New Roman","宋体","新宋体";font-size:12px;color:#0000FF;}
.BlueG14 {font-family:"Georgia","Arial","Times New Roman","宋体","新宋体";font-size:14px;color:#0000FF;}
.BlueG16 {font-family:"Georgia","Arial","Times New Roman","宋体","新宋体";font-size:16px;color:#0000FF;}
.BlueG18 {font-family:"Georgia","Arial","Times New Roman","宋体","新宋体";font-size:18px;color:#0000FF;}
.BlueS12 {font-family:"宋体","新宋体";font-size:12px;color:#0000FF;}
.BlueS14 {font-family:"宋体","新宋体";font-size:14px;color:#0000FF;}
.BlueS16 {font-family:"宋体","新宋体";font-size:16px;color:#0000FF;}
.BlueS18 {font-family:"宋体","新宋体";font-size:18px;color:#0000FF;}
.BlueV12 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:12px;color:#0000FF;}
.BlueV14 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:14px;color:#0000FF;}
.BlueV16 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:16px;color:#0000FF;}
.BlueV18 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:18px;color:#0000FF;}

.GrayS12 {font-family:"宋体","新宋体";font-size:12px;color:#808080;}
.GrayS14 {font-family:"宋体","新宋体";font-size:14px;color:#808080;}
.GrayS16 {font-family:"宋体","新宋体";font-size:16px;color:#808080;}
.GrayV12 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:12px;color:#808080;}
.GrayV14 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:14px;color:#808080;}
.GrayV16 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:16px;color:#808080;}

.LightgrayS12 {font-family:"宋体","新宋体";font-size:12px;color:#0000A0;}
.LightgrayS14 {font-family:"宋体","新宋体";font-size:14px;color:#0000A0;}
.LightgrayS16 {font-family:"宋体","新宋体";font-size:16px;color:#0000A0;}
.LightgrayV12 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:12px;color:#0000A0;}
.LightgrayV14 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:14px;color:#0000A0;}
.LightgrayV16 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:16px;color:#0000A0;}

.MaroonG12 {font-family:"Georgia","Arial","Times New Roman","宋体","新宋体";font-size:12px;color:#840000;}
.MaroonG14 {font-family:"Georgia","Arial","Times New Roman","宋体","新宋体";font-size:14px;color:#840000;}
.MaroonG16 {font-family:"Georgia","Arial","Times New Roman","宋体","新宋体";font-size:16px;color:#840000;}
.MaroonS12 {font-family:"宋体","新宋体";font-size:12px;color:#840000;}
.MaroonS14 {font-family:"宋体","新宋体";font-size:14px;color:#840000;}
.MaroonS16 {font-family:"宋体","新宋体";font-size:16px;color:#840000;}
.MaroonV12 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:12px;color:#840000;}
.MaroonV14 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:14px;color:#840000;}
.MaroonV16 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:16px;color:#840000;}

.NavyS12 {font-family:"宋体","新宋体";font-size:12px;color:#000080;}
.NavyS14 {font-family:"宋体","新宋体";font-size:14px;color:#000080;}
.NavyS16 {font-family:"宋体","新宋体";font-size:16px;color:#000080;}
.NavyV12 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:12px;color:#000080;}
.NavyV14 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:14px;color:#000080;}
.NavyV16 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:16px;color:#000080;}

.RedG12 {font-family:"Georgia","Arial","Times New Roman","宋体","新宋体";font-size:12px;color:#FF0000;}
.RedG14 {font-family:"Georgia","Arial","Times New Roman","宋体","新宋体";font-size:14px;color:#FF0000;}
.RedG16 {font-family:"Georgia","Arial","Times New Roman","宋体","新宋体";font-size:16px;color:#FF0000;}
.RedG18 {font-family:"Georgia","Arial","Times New Roman","宋体","新宋体";font-size:18px;color:#FF0000;}
.RedS12 {font-family:"宋体","新宋体";font-size:12px;color:#FF0000;}
.RedS14 {font-family:"宋体","新宋体";font-size:14px;color:#FF0000;}
.RedS16 {font-family:"宋体","新宋体";font-size:16px;color:#FF0000;}
.RedS18 {font-family:"宋体","新宋体";font-size:18px;color:#FF0000;}
.RedV12 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:12px;color:#FF0000;}
.RedV14 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:14px;color:#FF0000;}
.RedV16 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:16px;color:#FF0000;}
.RedV18 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:18px;color:#FF0000;}

.TealS12 {font-family:"宋体","新宋体";font-size:12px;color:#008080;}
.TealS14 {font-family:"宋体","新宋体";font-size:14px;color:#008080;}
.TealS16 {font-family:"宋体","新宋体";font-size:16px;color:#008080;}
.TealV12 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:12px;color:#008080;}
.TealV14 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:14px;color:#008080;}
.TealV16 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:16px;color:#008080;}

.SilverG12 {font-family:"Georgia","Arial","Times New Roman","宋体","新宋体";font-size:12px;color:#C0C0C0;}
.SilverG14 {font-family:"Georgia","Arial","Times New Roman","宋体","新宋体";font-size:14px;color:#C0C0C0;}
.SilverG16 {font-family:"Georgia","Arial","Times New Roman","宋体","新宋体";font-size:16px;color:#C0C0C0;}
.SilverG18 {font-family:"Georgia","Arial","Times New Roman","宋体","新宋体";font-size:18px;color:#C0C0C0;}

.SilverS12 {font-family:"宋体","新宋体";font-size:12px;color:#C0C0C0;}
.SilverS14 {font-family:"宋体","新宋体";font-size:14px;color:#C0C0C0;}
.SilverS16 {font-family:"宋体","新宋体";font-size:16px;color:#C0C0C0;}
.SilverS18 {font-family:"宋体","新宋体";font-size:18px;color:#C0C0C0;}
.SilverV12 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:12px;color:#C0C0C0;}
.SilverV14 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:14px;color:#C0C0C0;}
.SilverV16 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:16px;color:#C0C0C0;}
.SilverV18 {font-family:"Verdana","Arial","Times New Roman","宋体","新宋体";font-size:18px;color:#C0C0C0;}
.SilverW12 {font-family:"Wingdings";font-size:12px;color:#C0C0C0;}
.SilverW14 {font-family:"Wingdings";font-size:14px;color:#C0C0C0;}
.SilverW16 {font-family:"Wingdings";font-size:16px;color:#C0C0C0;}
.SilverW18 {font-family:"Wingdings";font-size:18px;color:#C0C0C0;}

.Style_Menu {font-family:微软雅黑,黑体,宋体;font-size:16px;color:#000000;font-weight:;}
.Style_Title_Form {font-family:微软雅黑,黑体,宋体;font-size:16px;color:#000000;font-weight:;}
.Style_Title_List {font-family:微软雅黑,黑体,宋体;font-size:16px;color:#000000;font-weight:;}
.Style_Title_Message {font-family:微软雅黑,黑体,宋体;font-size:16px;color:#000000;font-weight:;}

.Style_Table_Edit_Whole {width:960px;height:auto;background:#CEEFE7;}
.Style_Table_Edit_Title {width:auto;height:34px;background:transparent url('Guestbook.png') repeat-x scroll 0px -792px;}
.Style_Table_Edit_Distance {width:auto;height:6px;background:#FFFFFF;}
.Style_Table_Edit_Name {width:90px;height:36px;background:#FFFFFF;}
.Style_Table_Edit_Form {width:387px;height:36px;background:#FFFFFF;}
.Style_Table_Edit_Note {width:auto;height:36px;background:#FFFFFF;}
.Style_Table_Edit_Operate {width:auto;height:44px;background:#FFFFFF;padding:0px 0px 0px 0px;vertical-align:middle;}

.Style_Table_List_Whole {width:960px;height:auto;background:#CEEFE7;}
.Style_Table_List_Title {width:auto;height:34px;background:transparent url('Guestbook.png') repeat-x scroll 0px -792px;}
.Style_Table_List_Distance {width:auto;height:6px;background:#FFFFFF;}
.Style_Table_List_Name {width:auto;height:36px;background:#EAFBF5;}
.Style_Table_List_Content {width:auto;height:36px;}
.Style_Table_List_Operate {width:auto;height:44px;background:#FFFFFF;vertical-align:middle;}

.Style_Pagination_Admin_BackNext {width:100%;text-align:left;}
.Style_Pagination_Admin_BackNext a{height:24px;line-height:24px;display:inline-block;border:1px solid #A7A6AA;background:#F0F0F0;padding:0px 8px 0px 8px;font-family:宋体;font-size:12px;color:#000000;text-decoration:none;}
.Style_Pagination_Admin_BackNext a:hover{border:1px solid #A7A6AA;background:#A7A6AA;font-family:宋体;font-size:12px;color:#FFFFFF;text-decoration:none;}
.Style_Pagination_Admin_BackNext span{height:24px;line-height:24px;display:inline-block;border:1px solid #C0C0C0;background:#F9F9F9;padding:0px 8px 0px 8px;font-family:宋体;font-size:12px;color:#C0C0C0;text-decoration:none;}
.Style_Pagination_Admin_Number {width:100%;text-align:right;}
.Style_Pagination_Admin_Number a{height:24px;line-height:24px;display:inline-block;border:1px solid #A7A6AA;background:#F0F0F0;padding:0px 6px 0px 6px;font-family:Verdana,Arial,宋体;font-size:12px;color:#000000;text-decoration:none;}
.Style_Pagination_Admin_Number a:hover{border:1px solid #A7A6AA;background:#A7A6AA;font-family:Verdana,Arial,宋体;font-size:12px;color:#FFFFFF;text-decoration:none;}
.Style_Pagination_Admin_Number a.Present{border:1px solid #BAB9BD;background:#D0CFD1;font-family:Verdana,Arial,宋体;font-size:12px;color:#FFFFFF;text-decoration:none;}
.Style_Pagination_Admin_InputText {width:36px;height:14px;border-width:1px 1px 1px 1px;border-color:#C0C0C0 #C0C0C0 #C0C0C0 #C0C0C0;background:#FFFFFF;font-family:Verdana,Arial,宋体;font-size:12px;color:#000000;vertical-align:middle;text-align:center;}
.Style_Pagination_Admin_Submit {width:0px;height:18px;border:0px solid #FFFFFF;background:#FFFFFF;padding:0px 0px 0px 0px;vertical-align:bottom;}

.Style_Button_Add {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px 0px;cursor:pointer;}
.Style_Button_Back {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -24px;cursor:pointer;}
.Style_Button_Cancel {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -48px;cursor:pointer;}
.Style_Button_Confirm {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -72px;cursor:pointer;}
.Style_Button_Copy {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -96px;cursor:pointer;}
.Style_Button_Create {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -120px;cursor:pointer;}
.Style_Button_Cut {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -144px;cursor:pointer;}
.Style_Button_Delete {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -168px;cursor:pointer;}
.Style_Button_Disable {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -192px;cursor:pointer;}
.Style_Button_Empty {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -216px;cursor:pointer;}
.Style_Button_Enabled {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -240px;cursor:pointer;}
.Style_Button_Execute {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -264px;cursor:pointer;}
.Style_Button_Goto {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -286px;cursor:pointer;}
.Style_Button_Help {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -312px;cursor:pointer;}
.Style_Button_Hidden {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -336px;cursor:pointer;}
.Style_Button_HTML {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -360px;cursor:pointer;}
.Style_Button_List {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -384px;cursor:pointer;}
.Style_Button_Modify {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -408px;cursor:pointer;}
.Style_Button_Option {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -432px;cursor:pointer;}
.Style_Button_Paste {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -456px;cursor:pointer;}
.Style_Button_Preview {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -480px;cursor:pointer;}
.Style_Button_Read {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -504px;cursor:pointer;}
.Style_Button_Reload {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -528px;cursor:pointer;}
.Style_Button_Reset {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -552px;cursor:pointer;}
.Style_Button_Save {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -576px;cursor:pointer;}
.Style_Button_Search {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -600px;cursor:pointer;}
.Style_Button_Select_All {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -624px;cursor:pointer;}
.Style_Button_Select_Clear {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -648px;cursor:pointer;}
.Style_Button_Select_Reverse {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -672px;cursor:pointer;}
.Style_Button_Submit {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -696px;cursor:pointer;}
.Style_Button_Show {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -720px;cursor:pointer;}
.Style_Button_Unread {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -744px;cursor:pointer;}
.Style_Button_Upload {width:64px;height:24px;border:0px;background:transparent url('Guestbook.png') no-repeat scroll 0px -768px;cursor:pointer;}

.Style_Iframe {width:100%;height:100%;}
.Style_InputCheckbox {width:15px;height:15px;border-width:1px 1px 1px 1px;border-color:#C0C0C0 #C0C0C0 #C0C0C0 #C0C0C0;vertical-align:middle;}
.Style_InputRadio {width:16px;height:16px;border-width:1px 1px 1px 1px;border-color:#C0C0C0 #C0C0C0 #C0C0C0 #C0C0C0;vertical-align:middle;}
.Style_InputText {width:846px;height:24px;border-width:1px 1px 1px 1px;border-color:#C0C0C0 #C0C0C0 #C0C0C0 #C0C0C0;background:#FFFFFF;padding:0px 0px 0px 0px;font-family:Verdana,Arial,宋体;font-size:14px;color:#000000;}
.Style_Select {width:256px;height:26px;border-width:1px 1px 1px 1px;border-color:#C0C0C0 #C0C0C0 #C0C0C0 #C0C0C0;background:#FFFFFF;font-family:Verdana,Arial,宋体;font-size:14px;color:#000000;vertical-align:middle;}
.Style_Textarea {width:340px;height:46px;border-width:1px 1px 1px 1px;border-color:#C0C0C0 #C0C0C0 #C0C0C0 #C0C0C0;background:#FFFFFF;font-family:Verdana,Arial,宋体;font-size:14px;color:#000000;}

.Style_Space_Distance {font-family:Times New Roman,Verdana,Arial,宋体;font-size:6px;}
.Style_DIV_Distance {width:100%;height:auto;border-width:0px 0px 0px 0px;margin:0px 0px 0px 0px;padding:0px 0px 0px 0px;}
//-->
</style>
<script type="text/javascript">
<!--
if (top.location!=self.location) {
	top.location=self.location;
	}

function GuReturnElement(ElementName) {
	if (document.getElementById) {
		return document.getElementById(ElementName);
		}
	else {
		if (document.all) {
			return document.all[ElementName];
			}
		else {
			if (document.layers) {
				return document.layers[ElementName];
				}
			else {
				return (ElementName);
				}
			}
		}
	}

function GuElementCheckedAll(ElementName) {
	if ((ElementName!=null)&&(ElementName!="")) {
		for (var i=0;i<ElementName.length;i++) {
			ElementName[i].checked=true;
			}
		}
	}

function GuElementCheckedReverse(ElementName) {
	if ((ElementName!=null)&&(ElementName!="")) {
		for (var i=0;i<ElementName.length;i++) {
			ElementName[i].checked=!ElementName[i].checked;
			}
		}
	}

function GuElementCheckedClear(ElementName) {
	if ((ElementName!=null)&&(ElementName!="")) {
		for (var i=0;i<ElementName.length;i++) {
			ElementName[i].checked=false;
			}
		}
	}

function GuCopyClipboard(ElementName) {
	if ((ElementName!=null)&&(ElementName!="")) {
		ElementName.select();
		document.execCommand('Copy');
		}
	}

function GuPasteClipboard(ElementName) {
	if ((ElementName!=null)&&(ElementName!="")) {
		ElementName.focus();
		document.execCommand('Paste');
		}
	}

function GuElementStyleDisplayBAB1or(ElementNA,ElementNB,ElementNC,ElementND,ElementNE,ElementNF,ElementNG) {
	if ((ElementNA!=null)&&(ElementNB!=null)) {
		if (GuReturnElement(ElementNA).style.display=="") {
			GuReturnElement(ElementNA).style.display="none";
			GuReturnElement(ElementNB).style.display="";
			if ((ElementNE!=null)&&(ElementNE!="")) {
				GuReturnElement(ElementNE).innerHTML=ElementNF;
				}
			}
		else {
			GuReturnElement(ElementNA).style.display="";
			GuReturnElement(ElementNB).style.display="none";
			if ((ElementNC!=null)&&(ElementNC!="")) {
				GuReturnElement(ElementNC).value=ElementND.value;
				}
			if ((ElementNE!=null)&&(ElementNE!="")) {
				GuReturnElement(ElementNE).innerHTML=ElementNG;
				}
			}
		}
	}

function GuSubmitConfirm(ConfirmMessage,FormName,FormAction,ElementNA,ElementNB) {
	if ((FormName!=null)&&(FormName!="")) {
		var ElementCA=0;
		var ElementCB=0;
		if ((ElementNA!=null)&&(ElementNA!="")) {
			for (var i=0;i<ElementNA.length;i++) {
				if (ElementNA[i].checked==true) {
					ElementCA++;
					}
				}
			}
		if ((ElementNB!=null)&&(ElementNB!="")) {
			for (var i=0;i<ElementNB.length;i++) {
				if (ElementNB[i].checked==true) {
					ElementCB++;
					}
				}
			}
		if ((ElementCA>0)||(ElementCB>0)) {
			if (confirm(ConfirmMessage)) {
				FormName.action=FormAction;
				FormName.submit();
				}
			}
		else {
			FormName.action=FormAction;
			FormName.submit();
			}
		}
	}

function Guestbook_Login_Check(FormName) {
	if (FormName.GuAccount.value=="") {
		FormName.GuAccount.focus();
		FormName.GuAccount.style.backgroundColor="#FFECEC";
		GuReturnElement('JSinnerHTML_Login').innerHTML="<span style='color:#840000;'>没有填写管理账号。</span>";
		return (false);
		}
	if (FormName.GuPassword.value=="") {
		FormName.GuPassword.focus();
		FormName.GuPassword.style.backgroundColor="#FFECEC";
		GuReturnElement('JSinnerHTML_Login').innerHTML="<span style='color:#840000;'>没有填写登录密码。</span>";
		return (false);
		}
	if (GuReturnElement('GuCaptchaForm')!=null) {
		if (FormName.GuCaptchaForm.value=="") {
			FormName.GuCaptchaForm.focus();
			FormName.GuCaptchaForm.style.backgroundColor="#FFECEC";
			GuReturnElement('JSinnerHTML_Login').innerHTML="<span style='color:#840000;'>没有填写验证码。</span>";
			return (false);
			}
		}
	return (true);
	}

function Guestbook_Form_Check(FormName) {
	if ((FormName.GB_UserEMail.value.length>=1)&&((FormName.GB_UserEMail.value.length<=5)||(FormName.GB_UserEMail.value.indexOf("@")==-1)||(FormName.GB_UserEMail.value.indexOf(".")==-1))) {
		FormName.GB_UserEMail.focus();
		FormName.GB_UserEMail.style.backgroundColor="#FFECEC";
		GuReturnElement('JSinnerHTML_AddModify').innerHTML="<span style='color:#840000;'>没有填写正确的电子邮箱。</span>";
		return (false);
		}
	if ((FormName.GB_UserMobilePhone.value.length>=1)&&(FormName.GB_UserMobilePhone.value.length<=10)) {
		FormName.GB_UserMobilePhone.focus();
		FormName.GB_UserMobilePhone.style.backgroundColor="#FFECEC";
		GuReturnElement('JSinnerHTML_AddModify').innerHTML="<span style='color:#840000;'>没有填写正确的手机号码。</span>";
		return (false);
		}
	if ((FormName.GB_UserIMQQ.value.length>=1)&&(FormName.GB_UserIMQQ.value.length<=4)) {
		FormName.GB_UserIMQQ.focus();
		FormName.GB_UserIMQQ.style.backgroundColor="#FFECEC";
		GuReturnElement('JSinnerHTML_AddModify').innerHTML="<span style='color:#840000;'>没有填写正确的<span style='font-family:Verdana,Times New Roman,Tahoma;'>QQ</span>账号。</span>";
		return (false);
		}
	if (FormName.GB_Content.value=="") {
		FormName.GB_Content.focus();
		FormName.GB_Content.style.backgroundColor="#FFECEC";
		GuReturnElement('JSinnerHTML_AddModify').innerHTML="<span style='color:#840000;'>没有填写留言内容。</span>";
		return (false);
		}
	if (FormName.GB_Content.value.length>=<%=GuestbookContentLength%>+1) {
		FormName.GB_Content.focus();
		FormName.GB_Content.style.backgroundColor="#FFECEC";
		GuReturnElement('JSinnerHTML_AddModify').innerHTML="<span style='color:#840000;'>留言内容字符超过最大限制，最多<span style='font-family:Verdana,Times New Roman,Tahoma;'><%=GuestbookContentLength%></span>个。</span>";
		return (false);
		}
	if (GuReturnElement('GuCaptchaForm')!=null) {
		if (FormName.GuCaptchaForm.value=="") {
			FormName.GuCaptchaForm.focus();
			FormName.GuCaptchaForm.style.backgroundColor="#FFECEC";
			GuReturnElement('JSinnerHTML_AddModify').innerHTML="<span style='color:#840000;'>没有填写验证码。</span>";
			return (false);
			}
		}
	return (true);
	}

function Guestbook_Reply_Check(ElementName) {
	if (GuReturnElement(ElementName).value=="") {
		GuReturnElement(ElementName).value="网站管理员回复：";
		}
	}

function Guestbook_Reply_Submit(ResourceID) {
	if ((ResourceID!=null)&&(ResourceID!="")) {
		document.forms['Guestbook_Reply'].action="?Command=Reply&Type=Guestbook&Status=<%=trim(request.querystring("Status"))%>&Keyword=<%=trim(request.querystring("Keyword"))%>&Page=<%=trim(request.querystring("Page"))%>&ID="+ResourceID;
		document.forms['Guestbook_Reply'].GB_Reply.value=GuReturnElement('GB_Reply_'+ResourceID).value;
		document.forms['Guestbook_Reply'].submit();
		}
	}

function Guestbook_AddModify_Option(OptionKeywordOrder,OptionNameStyle) {
	if (GuReturnElement(OptionKeywordOrder).style.display=="") {
		GuReturnElement(OptionKeywordOrder).style.display="none";
		GuReturnElement(OptionNameStyle).style.display="";
		}
	else {
		GuReturnElement(OptionKeywordOrder).style.display="";
		GuReturnElement(OptionNameStyle).style.display="none";
		}
	}

function Guestbook_Search_Check(FormName,FormType) {
	if (FormType=="Guestbook") {
		if ((FormName.Keyword.value=="")||(FormName.Keyword.value=="没有输入关键字词")) {
			FormName.Keyword.focus();
			FormName.Keyword.value="没有输入关键字词";
			FormName.Keyword.style.backgroundColor="#FFECEC";
			return (false);
			}
		FormName.action="<%=WebpageName%>?Command=Search&Type=Guestbook&Keyword="+FormName.Keyword.value;
		FormName.submit();
		}
	}

function Administrator_AddModify_Check(FormName) {
	if (FormName.AM_Account.value=="") {
		FormName.AM_Account.focus();
		FormName.AM_Account.style.backgroundColor="#FFECEC";
		GuReturnElement('JSinnerHTML_Account').innerHTML="<span style='color:#840000;'>没有填写管理账号。</span>";
		return (false);
		}
	if (FormName.AM_Password.value=="") {
		FormName.AM_Password.focus();
		FormName.AM_Password.style.backgroundColor="#FFECEC";
		GuReturnElement('JSinnerHTML_Account').innerHTML="<span style='color:#840000;'>没有填写当前账号的登录密码。</span>";
		return (false);
		}
	if (FormName.AM_PasswordNew.value!=FormName.AM_PasswordConfirm.value) {
		FormName.AM_PasswordNew.style.backgroundColor="#FFECEC";
		FormName.AM_PasswordConfirm.focus();
		FormName.AM_PasswordConfirm.style.backgroundColor="#FFECEC";
		GuReturnElement('JSinnerHTML_Account').innerHTML="<span style='color:#840000;'>两次输入的登录密码没有一致。</span>";
		return (false);
		}
	return (true);
	}
//-->
</script>
</head>

<body style="margin:0px 0px 0px 0px;background:#FFFFFF;text-align:center;">
<div align="center">

<%
function GuDateTimeStyle(DateTimeData)

	if trim(DateTimeData)="" or isDate(DateTimeData)=false then
		GuDateTimeStyle="BlackS14"
		exit function
	end if

	if DateValue(DateTimeData)=date() then
		GuDateTimeStyle="MaroonS14"
	elseif datediff("d",DateTimeData,now())<=3 then
		GuDateTimeStyle="LightgrayS14"
	elseif datediff("d",DateTimeData,now())<=7 then
		GuDateTimeStyle="TealS14"
	else
		GuDateTimeStyle="BlackS14"
	end if

end function


function GuReplaceAccount(StringData)

	if trim(StringData)="" or isNull(StringData)=true then
		GuReplaceAccount=""
		exit function
	end if

	GuReplaceAccount=trim(StringData)
	GuReplaceAccount=replace(GuReplaceAccount,"~","")
	GuReplaceAccount=replace(GuReplaceAccount,"`","")
	GuReplaceAccount=replace(GuReplaceAccount,"!","")
	GuReplaceAccount=replace(GuReplaceAccount,"@","")
	GuReplaceAccount=replace(GuReplaceAccount,"#","")
	GuReplaceAccount=replace(GuReplaceAccount,"$","")
	GuReplaceAccount=replace(GuReplaceAccount,"%","")
	GuReplaceAccount=replace(GuReplaceAccount,"^","")
	GuReplaceAccount=replace(GuReplaceAccount,"&","")
	GuReplaceAccount=replace(GuReplaceAccount,"*","")
	GuReplaceAccount=replace(GuReplaceAccount,"+","")
	GuReplaceAccount=replace(GuReplaceAccount,"=","")
	GuReplaceAccount=replace(GuReplaceAccount,"{","")
	GuReplaceAccount=replace(GuReplaceAccount,"[","")
	GuReplaceAccount=replace(GuReplaceAccount,"}","")
	GuReplaceAccount=replace(GuReplaceAccount,"]","")
	GuReplaceAccount=replace(GuReplaceAccount,":","")
	GuReplaceAccount=replace(GuReplaceAccount,";","")
	GuReplaceAccount=replace(GuReplaceAccount,"'","")
	GuReplaceAccount=replace(GuReplaceAccount,"""","")
	GuReplaceAccount=replace(GuReplaceAccount,"|","")
	GuReplaceAccount=replace(GuReplaceAccount,"\","")
	GuReplaceAccount=replace(GuReplaceAccount,"<","")
	GuReplaceAccount=replace(GuReplaceAccount,",","")
	GuReplaceAccount=replace(GuReplaceAccount,">","")
	GuReplaceAccount=replace(GuReplaceAccount,"?","")
	GuReplaceAccount=replace(GuReplaceAccount,"/","")

end function


function GuReplaceResourceName(StringData)

	if trim(StringData)="" or isNull(StringData)=true then
		GuReplaceResourceName=""
		exit function
	end if

	GuReplaceResourceName=trim(StringData)
	GuReplaceResourceName=replace(GuReplaceResourceName,"'","‘")
	GuReplaceResourceName=replace(GuReplaceResourceName,"""","“")

end function


function GuReplaceResourceOther(StringData)

	if trim(StringData)="" or isNull(StringData)=true then
		GuReplaceResourceOther=""
		exit function
	end if

	GuReplaceResourceOther=trim(StringData)
	GuReplaceResourceOther=replace(GuReplaceResourceOther,"~","～")
	GuReplaceResourceOther=replace(GuReplaceResourceOther,"`","｀")
	GuReplaceResourceOther=replace(GuReplaceResourceOther,"#","＃")
	GuReplaceResourceOther=replace(GuReplaceResourceOther,"$","§")
	GuReplaceResourceOther=replace(GuReplaceResourceOther,"^","∧")
	GuReplaceResourceOther=replace(GuReplaceResourceOther,"*","＊")
	GuReplaceResourceOther=replace(GuReplaceResourceOther,"{","｛")
	GuReplaceResourceOther=replace(GuReplaceResourceOther,"}","｝")
	GuReplaceResourceOther=replace(GuReplaceResourceOther,";","；")
	GuReplaceResourceOther=replace(GuReplaceResourceOther,"'","‘")
	GuReplaceResourceOther=replace(GuReplaceResourceOther,"""","“")
	GuReplaceResourceOther=replace(GuReplaceResourceOther,"\","＼")
	GuReplaceResourceOther=replace(GuReplaceResourceOther,"<","＜")
	GuReplaceResourceOther=replace(GuReplaceResourceOther,",","，")
	GuReplaceResourceOther=replace(GuReplaceResourceOther,">","＞")
	GuReplaceResourceOther=replace(GuReplaceResourceOther,"and","Ａnd")
	GuReplaceResourceOther=replace(GuReplaceResourceOther,"delete","Ｄelete")
	GuReplaceResourceOther=replace(GuReplaceResourceOther,"update","Ｕpdate")
	GuReplaceResourceOther=replace(GuReplaceResourceOther,"insert","Ｉnsert")
	GuReplaceResourceOther=replace(GuReplaceResourceOther,"select","Ｓelect")

end function


function GuReplaceSearchKeyword(StringData)

	if trim(StringData)="" or isNull(StringData)=true then
		GuReplaceSearchKeyword=""
		exit function
	end if

	GuReplaceSearchKeyword=trim(StringData)
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"~","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"`","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"!","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"@","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"#","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"$","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"%","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"^","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"&","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"*","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"+","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"=","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"{","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"[","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"}","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"]","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,":","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,";","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"'","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"""","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"|","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"\","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"<","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,",","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,">","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"?","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"/","n")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"and","Ａnd")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"delete","Ｄelete")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"update","ＵPdate")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"insert","Ｉnsert")
	GuReplaceSearchKeyword=replace(GuReplaceSearchKeyword,"select","Ｓelect")

end function


function GuResourceStatusText(ResourceStatusCode)

	if trim(ResourceStatusCode)="" or isNull(ResourceStatusCode)=true then
		GuResourceStatusText="--"
		exit function
	end if

	select case ResourceStatusCode
	case "Approve"
		GuResourceStatusText="审核通过"
	case "Approving"
		GuResourceStatusText="审核中"
	case "Delete"
		GuResourceStatusText="删除[注销]"
	case "Disabled"
		GuResourceStatusText="禁用"
	case "Enabled"
		GuResourceStatusText="启用"
	case "Expire"
		GuResourceStatusText="过期"
	case "Hidden"
		GuResourceStatusText="隐藏"
	case "Normal"
		GuResourceStatusText="正常"
	case "Read"
		GuResourceStatusText="已读"
	case "Show"
		GuResourceStatusText="显示"
	case "Wait"
		GuResourceStatusText="等待"
	case "Unread"
		GuResourceStatusText="未读"
	case "Used"
		GuResourceStatusText="已用"
	case else
		GuResourceStatusText="--"
	end select

	GuResourceStatusText=GuResourceStatusText

end function


Private Const BITS_TO_A_BYTE = 8
Private Const BYTES_TO_A_WORD = 4
Private Const BITS_TO_A_WORD = 32

Private m_lOnBits(30)
Private m_l2Power(30)

Private function LShift(lValue, iShiftBits)
	If iShiftBits = 0 Then
		LShift = lValue
		Exit Function
	ElseIf iShiftBits = 31 Then
		If lValue And 1 Then
			LShift = &H80000000
		Else
			LShift = 0
		End If
		Exit Function
	ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
		Err.Raise 6
	End If

	If (lValue And m_l2Power(31 - iShiftBits)) Then
		LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
	Else
		LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
	End If
end function

Private function RShift(lValue, iShiftBits)
	If iShiftBits = 0 Then
		RShift = lValue
		Exit Function
	ElseIf iShiftBits = 31 Then
		If lValue And &H80000000 Then
			RShift = 1
		Else
			RShift = 0
		End If
		Exit Function
	ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
		Err.Raise 6
	End If
	
	RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)

	If (lValue And &H80000000) Then
		RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
	End If
end function

Private function Rotateleft(lValue, iShiftBits)
	RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
end function

Private function AddUnsigned(lX, lY)
	dim lX4
	dim lY4
	dim lX8
	dim lY8
	dim lResult
 
	lX8 = lX And &H80000000
	lY8 = lY And &H80000000
	lX4 = lX And &H40000000
	lY4 = lY And &H40000000
 
	lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
 
	If lX4 And lY4 Then
		lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
	ElseIf lX4 Or lY4 Then
		If lResult And &H40000000 Then
			lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
		Else
			lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
		End If
	Else
		lResult = lResult Xor lX8 Xor lY8
	End If
 
	AddUnsigned = lResult
end function

Private function md5_F(x, y, z)
	md5_F = (x And y) Or ((Not x) And z)
end function

Private function md5_G(x, y, z)
	md5_G = (x And z) Or (y And (Not z))
end function

Private function md5_H(x, y, z)
	md5_H = (x Xor y Xor z)
end function

Private function md5_I(x, y, z)
	md5_I = (y Xor (x Or (Not z)))
end function

Private sub md5_FF(a, b, c, d, x, s, ac)
	a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_F(b, c, d), x), ac))
	a = Rotateleft(a, s)
	a = AddUnsigned(a, b)
end sub

Private sub md5_GG(a, b, c, d, x, s, ac)
	a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_G(b, c, d), x), ac))
	a = Rotateleft(a, s)
	a = AddUnsigned(a, b)
end sub

Private sub md5_HH(a, b, c, d, x, s, ac)
	a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_H(b, c, d), x), ac))
	a = Rotateleft(a, s)
	a = AddUnsigned(a, b)
end sub

Private sub md5_II(a, b, c, d, x, s, ac)
	a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_I(b, c, d), x), ac))
	a = Rotateleft(a, s)
	a = AddUnsigned(a, b)
end sub

Private function ConvertToWordArray(sMessage)
	dim lMessageLength
	dim lNumberOfWords
	dim lWordArray()
	dim lBytePosition
	dim lByteCount
	dim lWordCount
	
	Const MODULUS_BITS = 512
	Const CONGRUENT_BITS = 448
	
	lMessageLength = len(sMessage)
	
	lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
	Redim lWordArray(lNumberOfWords - 1)
	
	lBytePosition = 0
	lByteCount = 0
	Do Until lByteCount >= lMessageLength
		lWordCount = lByteCount \ BYTES_TO_A_WORD
		lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
		lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
		lByteCount = lByteCount + 1
	loop

	lWordCount = lByteCount \ BYTES_TO_A_WORD
	lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE

	lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

	lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
	lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
	
	ConvertToWordArray = lWordArray
end function

Private function WordToHex(lValue)
	dim lByte
	dim lCount
	
	For lCount = 0 To 3
		lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
		WordToHex = WordToHex & right("0" & Hex(lByte), 2)
	Next
end function

Public function MD5(sMessage)

	m_lOnBits(0) = CLng(1)
	m_lOnBits(1) = CLng(3)
	m_lOnBits(2) = CLng(7)
	m_lOnBits(3) = CLng(15)
	m_lOnBits(4) = CLng(31)
	m_lOnBits(5) = CLng(63)
	m_lOnBits(6) = CLng(127)
	m_lOnBits(7) = CLng(255)
	m_lOnBits(8) = CLng(511)
	m_lOnBits(9) = CLng(1023)
	m_lOnBits(10) = CLng(2047)
	m_lOnBits(11) = CLng(4095)
	m_lOnBits(12) = CLng(8191)
	m_lOnBits(13) = CLng(16383)
	m_lOnBits(14) = CLng(32767)
	m_lOnBits(15) = CLng(65535)
	m_lOnBits(16) = CLng(131071)
	m_lOnBits(17) = CLng(262143)
	m_lOnBits(18) = CLng(524287)
	m_lOnBits(19) = CLng(1048575)
	m_lOnBits(20) = CLng(2097151)
	m_lOnBits(21) = CLng(4194303)
	m_lOnBits(22) = CLng(8388607)
	m_lOnBits(23) = CLng(16777215)
	m_lOnBits(24) = CLng(33554431)
	m_lOnBits(25) = CLng(67108863)
	m_lOnBits(26) = CLng(134217727)
	m_lOnBits(27) = CLng(268435455)
	m_lOnBits(28) = CLng(536870911)
	m_lOnBits(29) = CLng(1073741823)
	m_lOnBits(30) = CLng(2147483647)
	
	m_l2Power(0) = CLng(1)
	m_l2Power(1) = CLng(2)
	m_l2Power(2) = CLng(4)
	m_l2Power(3) = CLng(8)
	m_l2Power(4) = CLng(16)
	m_l2Power(5) = CLng(32)
	m_l2Power(6) = CLng(64)
	m_l2Power(7) = CLng(128)
	m_l2Power(8) = CLng(256)
	m_l2Power(9) = CLng(512)
	m_l2Power(10) = CLng(1024)
	m_l2Power(11) = CLng(2048)
	m_l2Power(12) = CLng(4096)
	m_l2Power(13) = CLng(8192)
	m_l2Power(14) = CLng(16384)
	m_l2Power(15) = CLng(32768)
	m_l2Power(16) = CLng(65536)
	m_l2Power(17) = CLng(131072)
	m_l2Power(18) = CLng(262144)
	m_l2Power(19) = CLng(524288)
	m_l2Power(20) = CLng(1048576)
	m_l2Power(21) = CLng(2097152)
	m_l2Power(22) = CLng(4194304)
	m_l2Power(23) = CLng(8388608)
	m_l2Power(24) = CLng(16777216)
	m_l2Power(25) = CLng(33554432)
	m_l2Power(26) = CLng(67108864)
	m_l2Power(27) = CLng(134217728)
	m_l2Power(28) = CLng(268435456)
	m_l2Power(29) = CLng(536870912)
	m_l2Power(30) = CLng(1073741824)

	dim x
	dim k
	dim AA
	dim BB
	dim CC
	dim DD
	dim a
	dim b
	dim c
	dim d
	
	Const S11 = 7
	Const S12 = 12
	Const S13 = 17
	Const S14 = 22
	Const S21 = 5
	Const S22 = 9
	Const S23 = 14
	Const S24 = 20
	Const S31 = 4
	Const S32 = 11
	Const S33 = 16
	Const S34 = 23
	Const S41 = 6
	Const S42 = 10
	Const S43 = 15
	Const S44 = 21

	x = ConvertToWordArray(sMessage)
	
	a = &H67452301
	b = &HEFCDAB89
	c = &H98BADCFE
	d = &H10325476

	for k=0 to ubound(x) Step 16
		AA = a
		BB = b
		CC = c
		DD = d
	
		md5_FF a, b, c, d, x(k + 0), S11, &HD76AA478
		md5_FF d, a, b, c, x(k + 1), S12, &HE8C7B756
		md5_FF c, d, a, b, x(k + 2), S13, &H242070DB
		md5_FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
		md5_FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
		md5_FF d, a, b, c, x(k + 5), S12, &H4787C62A
		md5_FF c, d, a, b, x(k + 6), S13, &HA8304613
		md5_FF b, c, d, a, x(k + 7), S14, &HFD469501
		md5_FF a, b, c, d, x(k + 8), S11, &H698098D8
		md5_FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
		md5_FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
		md5_FF b, c, d, a, x(k + 11), S14, &H895CD7BE
		md5_FF a, b, c, d, x(k + 12), S11, &H6B901122
		md5_FF d, a, b, c, x(k + 13), S12, &HFD987193
		md5_FF c, d, a, b, x(k + 14), S13, &HA679438E
		md5_FF b, c, d, a, x(k + 15), S14, &H49B40821
	
		md5_GG a, b, c, d, x(k + 1), S21, &HF61E2562
		md5_GG d, a, b, c, x(k + 6), S22, &HC040B340
		md5_GG c, d, a, b, x(k + 11), S23, &H265E5A51
		md5_GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
		md5_GG a, b, c, d, x(k + 5), S21, &HD62F105D
		md5_GG d, a, b, c, x(k + 10), S22, &H2441453
		md5_GG c, d, a, b, x(k + 15), S23, &HD8A1E681
		md5_GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
		md5_GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
		md5_GG d, a, b, c, x(k + 14), S22, &HC33707D6
		md5_GG c, d, a, b, x(k + 3), S23, &HF4D50D87
		md5_GG b, c, d, a, x(k + 8), S24, &H455A14ED
		md5_GG a, b, c, d, x(k + 13), S21, &HA9E3E905
		md5_GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
		md5_GG c, d, a, b, x(k + 7), S23, &H676F02D9
		md5_GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A
			
		md5_HH a, b, c, d, x(k + 5), S31, &HFFFA3942
		md5_HH d, a, b, c, x(k + 8), S32, &H8771F681
		md5_HH c, d, a, b, x(k + 11), S33, &H6D9D6122
		md5_HH b, c, d, a, x(k + 14), S34, &HFDE5380C
		md5_HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
		md5_HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
		md5_HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
		md5_HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
		md5_HH a, b, c, d, x(k + 13), S31, &H289B7EC6
		md5_HH d, a, b, c, x(k + 0), S32, &HEAA127FA
		md5_HH c, d, a, b, x(k + 3), S33, &HD4EF3085
		md5_HH b, c, d, a, x(k + 6), S34, &H4881D05
		md5_HH a, b, c, d, x(k + 9), S31, &HD9D4D039
		md5_HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
		md5_HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
		md5_HH b, c, d, a, x(k + 2), S34, &HC4AC5665
	
		md5_II a, b, c, d, x(k + 0), S41, &HF4292244
		md5_II d, a, b, c, x(k + 7), S42, &H432AFF97
		md5_II c, d, a, b, x(k + 14), S43, &HAB9423A7
		md5_II b, c, d, a, x(k + 5), S44, &HFC93A039
		md5_II a, b, c, d, x(k + 12), S41, &H655B59C3
		md5_II d, a, b, c, x(k + 3), S42, &H8F0CCC92
		md5_II c, d, a, b, x(k + 10), S43, &HFFEFF47D
		md5_II b, c, d, a, x(k + 1), S44, &H85845DD1
		md5_II a, b, c, d, x(k + 8), S41, &H6FA87E4F
		md5_II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
		md5_II c, d, a, b, x(k + 6), S43, &HA3014314
		md5_II b, c, d, a, x(k + 13), S44, &H4E0811A1
		md5_II a, b, c, d, x(k + 4), S41, &HF7537E82
		md5_II d, a, b, c, x(k + 11), S42, &HBD3AF235
		md5_II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
		md5_II b, c, d, a, x(k + 9), S44, &HEB86D391

		a = AddUnsigned(a, AA)
		b = AddUnsigned(b, BB)
		c = AddUnsigned(c, CC)
		d = AddUnsigned(d, DD)
	Next

	MD5=ucase(WordToHex(a)&WordToHex(b)&WordToHex(c)&WordToHex(d))
'	MD5_16=lcase(WordToHex(b)&WordToHex(c))
'	MD5_32=ucase(WordToHex(a)&WordToHex(b)&WordToHex(c)&WordToHex(d))

end function


set ACMA=Server.CreateObject("ADODB.Connection")
ACMA.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(GuDatabaseFileName)


select case trim(request.querystring("Command"))
case "Add"
	if trim(request.form("GB_Content"))="" then
		GuSystemMessageContent="<span style=""color:#840000;"">没有填写留言内容。</span>"
		call GuestbookMenu()
		call GuestbookAddForm()
	else
		GB_UserName=trim(request.form("GB_UserName"))
		GB_UserEMail=trim(request.form("GB_UserEMail"))
		GB_UserMobilePhone=trim(request.form("GB_UserMobilePhone"))
		GB_UserIMQQ=trim(request.form("GB_UserIMQQ"))
		GB_UserIMWeixin=""
		GB_UserIMYixin=""
		GB_UserIMMomo=""
		GB_Content=trim(request.form("GB_Content"))
		GB_Reply=""
		GB_Status="Hidden"
		GB_Type=""
		GB_UserName=GuReplaceResourceOther(GB_UserName)
		GB_UserEMail=GuReplaceResourceOther(GB_UserEMail)
		GB_UserMobilePhone=GuReplaceResourceOther(GB_UserMobilePhone)
		GB_UserIMQQ=GuReplaceResourceOther(GB_UserIMQQ)
		call GuestbookAddExecute()
	end if
case "Reply","Show","Hidden","Delete","Edit","Modify","LogDelete"
	if AdministratorID="" or AdministratorAccount="" then
		GuSystemMessageContent="<span style=""color:#840000;"">操作失败，没有登录或者已经超时。</span>"
		call GuestbookMenu()
		call AdministratorLoginForm()
	else
		set ARMC=Server.CreateObject("ADODB.RecordSet")
		SQL="select * from "&GuDatabaseTablePrefix&"_Administrator where AM_ID="&AdministratorID&" and AM_Account='"&AdministratorAccount&"' and AM_LoginLdentifying='"&AdministratorLoginLdentifying&"' and AM_Status='Enabled' and AM_DatabaseTablePrefix='"&GuDatabaseTablePrefix&"'"
		ARMC.Open SQL,ACMA,1,1
		if ARMC.bof and ARMC.eof then
			AdministratorID=""
			AdministratorAccount=""
			AdministratorPurview=""
		else
			AdministratorID=ARMC("AM_ID")
			AdministratorAccount=ARMC("AM_Account")
			AdministratorPurview=ucase(ARMC("AM_Purview"))
		end if
		ARMC.close
		set ARMC=nothing
		if AdministratorID="" or AdministratorAccount="" then
			GuSystemMessageContent="<span style=""color:#840000;"">操作失败，没有登录或者已经超时。</span>"
			call GuestbookMenu()
			call AdministratorLoginForm()
		else
			call GuestbookSelectType()
		end if
	end if
case "List","Search"
	call GuestbookMenu()
	call GuestbookList()
case "Login"
	call AdministratorLoginExecute()
case "Logon"
	call GuestbookMenu()
	call AdministratorLoginForm()
case "Logout"
	call AdministratorLogoutExecute()
case "New"
	call GuestbookMenu()
	call GuestbookAddForm()
case else
	call GuestbookMenu()
	call GuestbookAddForm()
end select


sub GuestbookSelectType()

	select case trim(request.querystring("Type"))
	case "Guestbook"
		select case trim(request.querystring("Command"))
		case "Reply"
			call GuestbookReplyExecute()
		case "Show"
			call GuestbookShowExecute()
		case "Hidden"
			call GuestbookHiddenExecute()
		case "Delete"
			call GuestbookDeleteExecute()
		case else
			call GuestbookMenu()
			call GuestbookList()
		end select
	case "Account"
		select case trim(request.querystring("Command"))
		case "Add","Modify"
			if len(trim(request.form("AM_Account")))>=1 then
			if trim(request.form("AM_Account"))=AdministratorAccount then
				AdministratorModifyCheck="Keep"
			else
				set ARAC=Server.CreateObject("ADODB.RecordSet")
				SQL="select * from "&GuDatabaseTablePrefix&"_Administrator where AM_ID<>"&AdministratorID&" and AM_Account='"&trim(request.form("AM_Account"))&"'"
				ARAC.open SQL,ACMA,1,1
				if ARAC.bof and ARAC.eof then
					AdministratorModifyCheck="Allow"
				else
					AdministratorModifyCheck="Repeat"
				end if
				ARAC.close
				set ARAC=nothing
			end if
			end if

			if len(trim(request.form("AM_Password")))>=1 then
				AM_Password=MD5(MD5(trim(request.form("AM_Password"))))
				set ARAC=Server.CreateObject("ADODB.RecordSet")
				SQL="select * from "&GuDatabaseTablePrefix&"_Administrator where AM_Account='"&AdministratorAccount&"' and AM_Password='"&AM_Password&"' and AM_Status='Enabled' and AM_DatabaseTablePrefix='"&GuDatabaseTablePrefix&"'"
				ARAC.open SQL,ACMA,1,1
				if ARAC.bof and ARAC.eof then
					AdministratorPasswordConfirm="Error"
				else
					AdministratorPasswordConfirm="Correct"
				end if
				ARAC.close
				set ARAC=nothing
			end if

			if trim(request.form("AM_Account"))="" then
				GuSystemMessageContent="<span style=""color:#840000;"">没有填写管理账号。</span>"
				call GuestbookMenu()
				call AdministratorModifyForm(AdministratorModifyBoolean)
				call AdministratorLogList()
			elseif trim(request.form("AM_Password"))="" then
				GuSystemMessageContent="<span style=""color:#840000;"">没有填写验证密码。</span>"
				call GuestbookMenu()
				call AdministratorModifyForm(AdministratorModifyBoolean)
				call AdministratorLogList()
			elseif AdministratorModifyCheck="Repeat" then
				GuSystemMessageContent="<span style=""color:#840000;"">管理账号已经存在。</span>"
				call GuestbookMenu()
				call AdministratorModifyForm(AdministratorModifyBoolean)
				call AdministratorLogList()
			elseif AdministratorPasswordConfirm="Error" then
				GuSystemMessageContent="<span style=""color:#840000;"">当前管理账号的验证密码错误。</span>"
				call GuestbookMenu()
				call AdministratorModifyForm(AdministratorModifyBoolean)
				call AdministratorLogList()
			elseif trim(request.form("AM_PasswordNew"))<>trim(request.form("AM_PasswordConfirm")) then
				GuSystemMessageContent="<span style=""color:#840000;"">两次输入新的登录密码没有一致。</span>"
				call GuestbookMenu()
				call AdministratorModifyForm(AdministratorModifyBoolean)
				call AdministratorLogList()
			else
				AM_Account=trim(request.form("AM_Account"))
				AM_Account=GuReplaceAccount(AM_Account)

				if len(trim(request.form("AM_PasswordNew")))>=1 then
					AdministratorPasswordModify="Yes"
					AM_Password=MD5(MD5(trim(request.form("AM_PasswordNew"))))
				end if

				if trim(request.querystring("Command"))="Add" then
					call AdministratorAddExecute()
				else
					call AdministratorModifyExecute()
				end if
			end if
		case "Edit"
			AdministratorModifyBoolean=true
			call GuestbookMenu()
			call AdministratorModifyForm(AdministratorModifyBoolean)
			call AdministratorLogList()
		case "LogDelete"
			call AdministratorLogDeleteExecute()
		case else
			call GuestbookMenu()
			call AdministratorModifyForm(AdministratorModifyBoolean)
			call AdministratorLogList()
		end select
	case else
		call GuestbookMenu()
		call AdministratorLoginForm()
	end select
end sub


sub AdministratorLoginExecute()

	dim AM_Account,AM_Password
	dim AdministratorLoginRestrictedNumber:AdministratorLoginRestrictedNumber=0

	if trim(request.form("GuAccount"))="" or trim(request.form("GuPassword"))="" then
		GuSystemMessageContent="<span style=""color:#840000;"">登录失败，帐号或者密码错误。</span>"
		call GuestbookMenu()
		call AdministratorLoginForm()
	else
		if AdministratorLoginRestrictedEnabled="1" then
			set ARIP=Server.CreateObject("ADODB.RecordSet")
			SQL="select * from "&GuDatabaseTablePrefix&"_IP where IP_Type='AdministratorLoginRestricted' and IP_Status='Enabled' and datediff('s',now(),IP_Expire)>=1 order by IP_ID desc"
			ARIP.open SQL,ACMA,1,1
			if ARIP.bof and ARIP.eof then
				AdministratorLoginRestrictedNumber=0
			else
				do while not ARIP.eof
					if AdministratorLoginRestrictedNumber>=1 then
						exit do
					end if
					if (ARIP("IP_Same")=1 and strcomp(split(ARIP("IP_Address"),".")(0),split(GuBrowserIP,".")(0),1)=0) or (ARIP("IP_Same")=2 and strcomp(split(ARIP("IP_Address"),".")(0)&split(ARIP("IP_Address"),".")(1),split(GuBrowserIP,".")(0)&split(GuBrowserIP,".")(1),1)=0) or (ARIP("IP_Same")=3 and strcomp(split(ARIP("IP_Address"),".")(0)&split(ARIP("IP_Address"),".")(1)&split(ARIP("IP_Address"),".")(2),split(GuBrowserIP,".")(0)&split(GuBrowserIP,".")(1)&split(GuBrowserIP,".")(2),1)=0) or (ARIP("IP_Same")=4 and strcomp(ARIP("IP_Address"),GuBrowserIP,1)=0) then
						AdministratorLoginRestrictedNumber=AdministratorLoginRestrictedNumber+1
					else
						AdministratorLoginRestrictedNumber=0
					end if
				ARIP.movenext
				loop
			end if
			ARIP.close
			set ARIP=nothing
		end if

		if AdministratorLoginRestrictedNumber>=1 then
			GuSystemMessageContent="<span style=""color:#840000;"">"&AdministratorLoginRestrictedMessage&"</span>"
			call GuestbookMenu()
			call AdministratorLoginForm()
		elseif AdministratorLoginCaptchaShow="1" and ucase(trim(request.form("GuCaptchaForm")))<>ucase(Request.Cookies("GuCaptchaCookies")) then
			GuSystemMessageContent="<span style=""color:#840000;"">登录失败，验证码错误。</span>"
			call GuestbookMenu()
			call AdministratorLoginForm()
		else
			AM_Account=lcase(trim(request.form("GuAccount")))
			AM_Account=GuReplaceAccount(AM_Account)
			AM_Password=MD5(MD5(trim(request.form("GuPassword"))))

			set ARAC=Server.CreateObject("ADODB.RecordSet")
			SQL="select * from "&GuDatabaseTablePrefix&"_Administrator where AM_Account='"&AM_Account&"' and AM_Password='"&AM_Password&"' and AM_Status='Enabled' and AM_DatabaseTablePrefix='"&GuDatabaseTablePrefix&"' order by AM_ID asc"
			ARAC.open SQL,ACMA,1,3
			if ARAC.bof and ARAC.eof then
				ACMA.execute("insert into "&GuDatabaseTablePrefix&"_Administrator_Log(AL_Code,AL_Note,AL_AddAdministrator,AL_AddIP,AL_AddDateTime) values('2718','登录失败，帐号或者密码错误。','"&trim(request.form("GuAccount"))&"','"&GuBrowserIP&"','"&now()&"')")
				if AdministratorLoginRestrictedEnabled="1" then
					dim AdministratorLoginDateTimeNew,AdministratorLoginDateTimePrior
					set ARIP=Server.CreateObject("ADODB.RecordSet")
					SQL="select top "&AdministratorLoginRestrictedMistake&" * from "&GuDatabaseTablePrefix&"_Administrator_Log where AL_Code=2718 and AL_AddIP='"&GuBrowserIP&"' order by AL_ID desc"
					ARIP.open SQL,ACMA,1,1
					if ARIP.bof and ARIP.eof then
						response.write ""
					else
						if ARIP.recordcount>=AdministratorLoginRestrictedMistake then
							AdministratorLoginDateTimeNew=ARIP("AL_AddDateTime")
							ARIP.MoveLast
							AdministratorLoginDateTimePrior=ARIP("AL_AddDateTime")
							if datediff("s",AdministratorLoginDateTimePrior,AdministratorLoginDateTimeNew)<=AdministratorLoginRestrictedDistance then
								ACMA.execute("insert into "&GuDatabaseTablePrefix&"_IP(IP_Address,IP_Same,IP_Note,IP_Expire,IP_Status,IP_Type,IP_AddAdministrator,IP_AddIP,IP_AddDateTime,IP_ModifyAdministrator,IP_ModifyIP,IP_ModifyDateTime) values('"&GuBrowserIP&"','4','','"&dateadd("n",AdministratorLoginRestrictedMinute,now())&"','Enabled','AdministratorLoginRestricted','"&trim(request.form("GuAccount"))&"','"&GuBrowserIP&"',now(),'"&trim(request.form("GuAccount"))&"','"&GuBrowserIP&"',now())")
							else
								response.write ""
							end if
						else
							response.write ""
						end if
					end if
					ARIP.close
					set ARIP=nothing
				end if
				GuSystemMessageContent="<span style=""color:#840000;"">登录失败，帐号或者密码错误。</span>"
				call GuestbookMenu()
				call AdministratorLoginForm()
			else
				ARAC("AM_LoginNumber")=ARAC("AM_LoginNumber")+1
				ARAC("AM_LoginLdentifying")=MD5(now())
				ARAC("AM_LastLoginIP")=GuBrowserIP
				ARAC("AM_LastLoginDateTime")=now()
				ARAC.Update

				AdministratorID=ARAC("AM_ID")
				AdministratorAccount=ARAC("AM_Account")
				AdministratorPurview=ucase(ARAC("AM_Purview"))

				if AdministratorLoginCookiesEnabled="1" then
					response.cookies("GUEEONGUESTBOOKADMINISTRATOR")("ID")=ARAC("AM_ID")
					response.cookies("GUEEONGUESTBOOKADMINISTRATOR")("ACCOUNT")=ARAC("AM_Account")
					response.cookies("GUEEONGUESTBOOKADMINISTRATOR")("LOGINLDENTIFYING")=ARAC("AM_LoginLdentifying")
					response.cookies("GUEEONGUESTBOOKADMINISTRATOR").expires=now()+AdministratorLoginTimeout/1440
				else
					session("GUEEONGUESTBOOKADMINISTRATORID")=ARAC("AM_ID")
					session("GUEEONGUESTBOOKADMINISTRATORACCOUNT")=ARAC("AM_Account")
					session("GUEEONGUESTBOOKADMINISTRATORLOGINLDENTIFYING")=ARAC("AM_LoginLdentifying")
					session.timeout=AdministratorLoginTimeout
				end if

				ACMA.execute("insert into "&GuDatabaseTablePrefix&"_Administrator_Log(AL_Code,AL_Note,AL_AddAdministrator,AL_AddIP,AL_AddDateTime) values('2720','登录成功。','"&trim(request.form("GuAccount"))&"','"&GuBrowserIP&"','"&now()&"')")
				ACMA.execute("delete from "&GuDatabaseTablePrefix&"_IP where IP_ID>=1 and IP_Type='AdministratorLoginRestricted' and datediff('n',IP_AddDateTime,now())>"&AdministratorLoginRestrictedMinute&"")

				call GuestbookMenu()
				call GuestbookList()

			end if
			ARAC.close
			set ARAC=nothing
		end if
	end if

end sub


sub AdministratorLogoutExecute()

	if len(AdministratorID)>=1 then
		set ARAC=Server.CreateObject("ADODB.RecordSet")
		SQL="select * from "&GuDatabaseTablePrefix&"_Administrator where AM_ID="&AdministratorID&" and AM_Status='Enabled' and AM_DatabaseTablePrefix='"&GuDatabaseTablePrefix&"'"
		ARAC.open SQL,ACMA,1,1
		if ARAC.bof and ARAC.eof then
			response.write ""
			response.end
		else
			ACMA.execute("insert into "&GuDatabaseTablePrefix&"_Administrator_Log(AL_Code,AL_Note,AL_AddAdministrator,AL_AddIP,AL_AddDateTime) values('2000','退出成功。','"&AdministratorAccount&"','"&GuBrowserIP&"','"&now()&"')")
		end if
		ARAC.close
		set ARAC=nothing
	end if

	if AdministratorLoginCookiesEnabled="1" then
		response.cookies("GUEEONGUESTBOOKADMINISTRATOR")("ID")=""
		response.cookies("GUEEONGUESTBOOKADMINISTRATOR")("ACCOUNT")=""
		response.cookies("GUEEONGUESTBOOKADMINISTRATOR")("LOGINLDENTIFYING")=""
		response.cookies("GUEEONGUESTBOOKADMINISTRATOR").expires=#2000-01-01#
	else
		session("GUEEONGUESTBOOKADMINISTRATORID")=""
		session("GUEEONGUESTBOOKADMINISTRATORACCOUNT")=""
		session("GUEEONGUESTBOOKADMINISTRATORLOGINLDENTIFYING")=""
	end if

	GuSystemMessageContent="<span style=""color:#0000FF;"">退出成功。</span>"
	call GuestbookMenu()
	call AdministratorLoginForm()

end sub


sub GuestbookAddExecute()

	set ARAC=Server.CreateObject("ADODB.RecordSet")
	SQL="select * from "&GuDatabaseTablePrefix&"_Guestbook where GB_AddIP='"&GuBrowserIP&"' order by GB_ID desc"
	ARAC.open SQL,ACMA,1,1
	if ARAC.bof and ARAC.eof then
		GuestbookAddDistanceCheck="Empty"
	else
		if datediff("s",ARAC("GB_AddDateTime"),now())>=GuestbookAddDistanceSecond then
			GuestbookAddDistanceCheck="Empty"
		else
			GuestbookAddDistanceCheck="Added"
		end if
	end if
	ARAC.close
	set ARAC=nothing

	if GuestbookAddDistanceCheck="Added" then
		GuSystemMessageContent="<span style=""color:#840000;"">添加失败，同一<span style=""font-family:Verdana;"">IP</span>地址间隔<span style=""font-family:Verdana;"">"&GuestbookAddDistanceSecond&"</span>秒才能再次添加留言。</span>"
	elseif len(GB_Content)>GuestbookContentLength then
		GuSystemMessageContent="<span style=""color:#840000;"">添加失败，留言内容字符超过最大限制，最多<span style=""font-family:Verdana;"">"&GuestbookContentLength&"</span>个。</span>"
	elseif GuestbookAddCaptchaShow="1" and ucase(trim(request.form("GuCaptchaForm")))<>ucase(Request.Cookies("GuCaptchaCookies")) then
		GuSystemMessageContent="<span style=""color:#840000;"">添加失败，验证码错误。</span>"
	else
		set ARAS=Server.CreateObject("ADODB.RecordSet")
		SQL="select * from "&GuDatabaseTablePrefix&"_Guestbook"
		ARAS.Open SQL,ACMA,1,3
		ARAS.AddNew
		ARAS("GB_UserName")=GB_UserName
		ARAS("GB_UserEMail")=GB_UserEMail
		ARAS("GB_UserMobilePhone")=GB_UserMobilePhone
		ARAS("GB_UserIMQQ")=GB_UserIMQQ
		ARAS("GB_UserIMWeixin")=GB_UserIMWeixin
		ARAS("GB_UserIMYixin")=GB_UserIMYixin
		ARAS("GB_UserIMMomo")=GB_UserIMMomo
		ARAS("GB_Content")=GB_Content
		ARAS("GB_Reply")=GB_Reply
		ARAS("GB_Status")=GB_Status
		ARAS("GB_Type")=GB_Type
		ARAS("GB_AddAdministrator")="[User]"
		ARAS("GB_AddIP")=GuBrowserIP
		ARAS("GB_AddDateTime")=now()
		ARAS("GB_ModifyAdministrator")="[User]"
		ARAS("GB_ModifyIP")=GuBrowserIP
		ARAS("GB_ModifyDateTime")=now()
		ARAS.update
		ARAS.close
		set ARAS=nothing

		if GuestbookAddEMailEnabled="1" then
			GuestbookSMTPBody="用户昵称："&GB_UserName&"&nbsp;&nbsp;电子邮箱："&GB_UserEMail&"&nbsp;&nbsp;手机号码："&GB_UserMobilePhone&"&nbsp;&nbsp;QQ账号："&GB_UserIMQQ&"&nbsp;&nbsp;IP地址："&GuBrowserIP&"<br />留言内容："&GB_Content&"&nbsp;&nbsp;"&now()
			dim JMail
			set JMail=Server.CreateObject("JMail.Message")
			JMail.Silent=true
			JMail.Logging=true
			JMail.Charset="gb2312"
			JMail.ContentType="text/html"
			JMail.Priority=3
			JMail.MailServerUsername=GuestbookSMTPUsername
			JMail.MailServerPassword=GuestbookSMTPPassword
			JMail.From=GuestbookSMTPUsername
			JMail.FromName=WebsiteName
			JMail.AddRecipient(GuestbookSMTPAddRecipient)
			JMail.Subject=GuestbookSMTPSubject
			JMail.Body=GuestbookSMTPBody
			JMail.Send(GuestbookSMTPServer)
			JMail.close()
			set JMail=nothing
		end if

		GuSystemMessageContent="<span style=""color:#0000FF;"">新的留言已添加完成。</span>"
	end if
	call GuestbookMenu()
	call GuestbookAddForm()

end sub


sub GuestbookReplyExecute()

	GB_Reply=trim(request.form("GB_Reply"))
	GB_Reply=GuReplaceResourceOther(GB_Reply)
	ACMA.execute("update "&GuDatabaseTablePrefix&"_Guestbook set GB_Reply='"&GB_Reply&"',GB_ReplyAdministrator='"&AdministratorAccount&"',GB_ReplyIP='"&GuBrowserIP&"',GB_ReplyDateTime=now() where GB_ID in ("&trim(request.querystring("ID"))&")")
	call GuestbookMenu()
	call GuestbookList()
	call GuestbookJSAName(trim(request.querystring("ID")))

end sub


sub GuestbookShowExecute()

	if trim(request.form("ResourceID"))="" then
		GuSystemMessageContent="<span style=""color:#840000;"">没有选择要设置显示的用户留言。</span>"
	else
		ACMA.execute("update "&GuDatabaseTablePrefix&"_Guestbook set GB_Status='Show',GB_ModifyAdministrator='"&AdministratorAccount&"',GB_ModifyIP='"&GuBrowserIP&"',GB_ModifyDateTime=now() where GB_ID in ("&trim(request.form("ResourceID"))&")")
		ACMA.execute("insert into "&GuDatabaseTablePrefix&"_Administrator_Log(AL_Code,AL_Note,AL_AddAdministrator,AL_AddIP,AL_AddDateTime) values('2320','用户留言设置显示已执行完成。','"&AdministratorAccount&"','"&GuBrowserIP&"','"&now()&"')")
		GuSystemMessageContent="<span style=""color:#0000FF;"">用户留言设置显示已执行完成。</span>"
	end if
	call GuestbookMenu()
	call GuestbookList()
	call GuestbookJSAName("GuestbookBottom")

end sub


sub GuestbookHiddenExecute()

	if trim(request.form("ResourceID"))="" then
		GuSystemMessageContent="<span style=""color:#840000;"">没有选择要设置隐藏的用户留言。</span>"
	else
		ACMA.execute("update "&GuDatabaseTablePrefix&"_Guestbook set GB_Status='Hidden',GB_ModifyAdministrator='"&AdministratorAccount&"',GB_ModifyIP='"&GuBrowserIP&"',GB_ModifyDateTime=now() where GB_ID in ("&trim(request.form("ResourceID"))&")")
		ACMA.execute("insert into "&GuDatabaseTablePrefix&"_Administrator_Log(AL_Code,AL_Note,AL_AddAdministrator,AL_AddIP,AL_AddDateTime) values('2320','用户留言设置隐藏已执行完成。','"&AdministratorAccount&"','"&GuBrowserIP&"','"&now()&"')")
		GuSystemMessageContent="<span style=""color:#0000FF;"">用户留言设置隐藏已执行完成。</span>"
	end if
	call GuestbookMenu()
	call GuestbookList()
	call GuestbookJSAName("GuestbookBottom")

end sub


sub GuestbookDeleteExecute()

	if trim(request.form("ResourceID"))="" then
		GuSystemMessageContent="<span style=""color:#840000;"">没有选择要删除的用户留言。</span>"
	else
		ACMA.Execute("delete from "&GuDatabaseTablePrefix&"_Guestbook where GB_ID in ("&trim(request.form("ResourceID"))&")")
		ACMA.execute("insert into "&GuDatabaseTablePrefix&"_Administrator_Log(AL_Code,AL_Note,AL_AddAdministrator,AL_AddIP,AL_AddDateTime) values('2320','用户留言已删除完成。','"&AdministratorAccount&"','"&GuBrowserIP&"','"&now()&"')")
		GuSystemMessageContent="<span style=""color:#0000FF;"">用户留言已删除完成。</span>"
	end if
	call GuestbookMenu()
	call GuestbookList()
	call GuestbookJSAName("GuestbookBottom")

end sub


sub AdministratorAddExecute()

	response.write ""

end sub


sub AdministratorModifyExecute()

	set ARMC=Server.CreateObject("ADODB.RecordSet")
	SQL="select * from "&GuDatabaseTablePrefix&"_Administrator where AM_Account='"&AdministratorAccount&"'"
	ARMC.open SQL,ACMA,1,1
	if ARMC.bof and ARMC.eof then
		GuSystemMessageContent="<span style=""color:#840000;"">操作失败，没有登录或者已经超时。</span>"
		call GuestbookMenu()
		call AdministratorLoginForm()
	else
		set ARMS=Server.CreateObject("ADODB.RecordSet")
		SQL="select * from "&GuDatabaseTablePrefix&"_Administrator where AM_ID="&AdministratorID&""
		ARMS.open SQL,ACMA,1,3

		if AdministratorModifyCheck="Allow" then
			ARMS("AM_Account")=AM_Account
		end if

		if AdministratorPasswordModify="Yes" then
			ARMS("AM_Password")=AM_Password
		end if

		ARMS("AM_ModifyIP")=GuBrowserIP
		ARMS("AM_ModifyDateTime")=now()
		ARMS.update
		ARMS.close
		set ARMS=nothing

		ACMA.execute("insert into "&GuDatabaseTablePrefix&"_Administrator_Log(AL_Code,AL_Note,AL_AddAdministrator,AL_AddIP,AL_AddDateTime) values('2220','管理账号（"&AM_Account&"）/密码已修改完成。','"&AdministratorAccount&"','"&GuBrowserIP&"','"&now()&"')")
		GuSystemMessageContent="<span style=""color:#0000FF;"">管理账号（<span style=""font-family:Verdana;"">"&AM_Account&"</span>）/密码已修改完成。</span>"
		call GuestbookMenu()
		call AdministratorModifyForm(AdministratorModifyBoolean)
		call AdministratorLogList()
	end if
	ARMC.close
	set ARMC=nothing

end sub


sub AdministratorLogDeleteExecute()

	if trim(request.form("GuResourceID"))="" then
		GuSystemMessageContent="<span style=""color:#840000;"">没有选择要删除的管理账号日志。</span>"
	else
		ACMA.execute("delete from "&GuDatabaseTablePrefix&"_Administrator_Log where AL_ID in ("&trim(request.form("GuResourceID"))&")")
		ACMA.execute("insert into "&GuDatabaseTablePrefix&"_Administrator_Log(AL_Code,AL_Note,AL_AddAdministrator,AL_AddIP,AL_AddDateTime) values('2320','管理账号日志已删除完成。','"&AdministratorAccount&"','"&GuBrowserIP&"','"&now()&"')")
		GuSystemMessageContent="<span style=""color:#0000FF;"">管理账号日志已删除完成。</span>"
	end if
	call GuestbookMenu()
	call AdministratorModifyForm(AdministratorModifyBoolean)
	call AdministratorLogList()

end sub
%>


<%sub GuestbookMenu()%>
<table border="0px" cellpadding="0px" cellspacing="0px" style="width:100%;height:auto;">
 <tr>
  <td align="center" style="width:auto;height:auto;background:#CEEFE7;"><table border="0px" cellpadding="0px" cellspacing="0px" style="width:960px;"><tr><td align="right" style="width:auto;height:60px;"><span class="GrayS16"><a href="<%=WebpageName%>?Command=New&Type=Guestbook" class="Style_Menu">添加留言</a> | <a href="<%=WebpageName%>?Command=List&Type=Guestbook" class="Style_Menu">留言列表</a> | <a href="?Command=Edit&Type=Account&ID=<%=AdministratorID%>" class="Style_Menu">账号</a> | <%if AdministratorID="" then%><a href="<%=WebpageName%>?Command=Logon" target="_top" class="Style_Menu">登录</a><%else%><a href="<%=WebpageName%>?Command=Logout" target="_top" class="Style_Menu">退出</a><%end if%></span>&nbsp;</td></tr></table></td>
 </tr>
</table>
<br />
<%end sub%>


<%sub AdministratorLoginForm()%>
<table border="0px" cellpadding="0px" cellspacing="1px" class="Style_Table_Edit_Whole">
<form method="post" name="Guestbook_Login" action="<%=WebpageName%>?Command=Login" onsubmit="javascript:return Guestbook_Login_Check(this);">
 <tr>
  <td colspan="4" align="center" class="Style_Table_Edit_Title"><span class="Style_Title_Form">留 言 管 理 登 录</span></td>
 </tr>
 <tr>
  <td colspan="4" class="Style_Table_Edit_Distance"></td>
 </tr>
 <tr>
  <td align="center" class="Style_Table_Edit_Name"><span class="BlackS14">管理账号</span></td>
  <td align="left" class="Style_Table_Edit_Form">&nbsp;<input type="text" id="GuAccount" name="GuAccount" maxlength="40" value="<%=trim(request.querystring("Account"))%>" class="Style_InputText" style="width:367px;" /></td>
  <td align="center" class="Style_Table_Edit_Name"><span class="BlackS14">登录密码</span></td>
  <td align="left" class="Style_Table_Edit_Form" style="width:auto;">&nbsp;<input type="password" id="GuPassword" name="GuPassword" maxlength="40" value="<%=trim(request.querystring("Password"))%>" class="Style_InputText" style="width:367px;" /></td>
 </tr>
 <tr>
  <td colspan="4" align="left" class="Style_Table_Edit_Operate"><table border="0px" cellpadding="0px" cellspacing="0px" style="width:100%;height:auto;"><tr><td align="left" style="width:98px;"></td><td align="left" style="width:160px;"><input type="submit" id="" name="Button_Confirm" value="" class="Style_Button_Confirm" />&nbsp;<input type="reset" id="" name="Button_Reset" value="" class="Style_Button_Reset" /></td><td style="width:140px;"><%if AdministratorLoginCaptchaShow="1" then%><input type="text" id="GuCaptchaForm" name="GuCaptchaForm" maxlength="4" value="" class="Style_InputText" style="width:44px;height:22px;text-align:center;" />&nbsp;<img id="" src="Captcha.asp" alt="验证码" title="单击即可刷新验证码" onclick="javascript:this.src='Captcha.asp?'+Math.random();" style="width:46px;height:14px;border:0px;vertical-align:middle;cursor:pointer;" /><%end if%></td><td style="width:auto;"><span id="JSinnerHTML_Login" class="BlackS14"><%=GuSystemMessageContent%></span></td></tr></table></td>
 </tr>
</form>
</table>
<br />

<script type="text/javascript">
<!--
	document.forms['Guestbook_Login'].GuAccount.focus();
//-->
</script>
<%end sub%>


<%sub GuestbookAddForm()%>
<table border="0px" cellpadding="0px" cellspacing="1px" class="Style_Table_Edit_Whole">
<form method="post" name="Guestbook_Form" action="?Command=Add&Type=Guestbook" onsubmit="javascript:return Guestbook_Form_Check(this);">
 <tr>
  <td colspan="4" align="center" class="Style_Table_Edit_Title"><span class="Style_Title_Form">用 户 留 言</span></td>
 </tr>
 <tr>
  <td colspan="4" class="Style_Table_Edit_Distance"></td>
 </tr>
 <tr>
  <td align="center" class="Style_Table_Edit_Name"><span class="BlackS14">用户昵称</span></td>
  <td align="left" class="Style_Table_Edit_Form">&nbsp;<input type="text" id="GB_UserName" name="GB_UserName" maxlength="40" value="" class="Style_InputText" style="width:367px;" /></td>
  <td align="center" class="Style_Table_Edit_Name"><span class="BlackS14">电子邮箱</span></td>
  <td align="left" class="Style_Table_Edit_Form" style="width:auto;">&nbsp;<input type="text" id="GB_UserEMail" name="GB_UserEMail" maxlength="40" value="" class="Style_InputText" style="width:367px;" /></td>
 </tr>
 <tr>
  <td align="center" class="Style_Table_Edit_Name"><span class="BlackS14">手机号码</span></td>
  <td align="left" class="Style_Table_Edit_Form">&nbsp;<input type="text" id="GB_UserMobilePhone" name="GB_UserMobilePhone" maxlength="40" value="" onafterpaste="javascript:this.value=this.value.replace(/\D/g,'');" onkeyup="javascript:this.value=this.value.replace(/\D/g,'');" class="Style_InputText" style="width:367px;" /></td>
  <td align="center" class="Style_Table_Edit_Name"><span class="BlackS14" style="font-family:Verdana,Times New Roman,Tahoma;">QQ</span><span class="BlackS14"></span>账号</span></td>
  <td align="left" class="Style_Table_Edit_Form" style="width:auto;">&nbsp;<input type="text" id="GB_UserIMQQ" name="GB_UserIMQQ" maxlength="40" value="" class="Style_InputText" style="width:367px;" /></td>
 </tr>
 <tr>
  <td align="center" class="Style_Table_Edit_Name" style="height:140px;vertical-align:top;"><table border="0px" cellpadding="0px" cellspacing="0px" style="width:100%;height:auto;"><tr><td align="center" style="width:auto;height:36px;"><span class="BlackS14">留言内容</span></td></tr><tr><td class="Style_Table_Edit_Whole" style="width:auto;height:1px;"></td></tr></table></td>
  <td colspan="3" align="left" class="Style_Table_Edit_Form" style="width:auto;padding:6px 0px 0px 0px;vertical-align:top;">&nbsp;<textarea id="GB_Content" name="GB_Content" rows="1" cols="1" class="Style_Textarea" style="width:844px;height:124px;"></textarea></td>
 </tr>
 <tr>
  <td colspan="4" align="left" class="Style_Table_Edit_Operate"><table border="0px" cellpadding="0px" cellspacing="0px" style="width:100%;height:auto;"><tr><td align="left" style="width:98px;"></td><td align="left" style="width:160px;"><input type="submit" id="" name="Button_Add" value="" class="Style_Button_Add" />&nbsp;<input type="reset" id="" name="Button_Reset" value="" class="Style_Button_Reset" /></td><td style="width:140px;"><%if GuestbookAddCaptchaShow="1" then%><input type="text" id="GuCaptchaForm" name="GuCaptchaForm" maxlength="4" value="" class="Style_InputText" style="width:44px;height:22px;text-align:center;" />&nbsp;<img id="" src="Captcha.asp" alt="验证码" title="单击即可刷新验证码" onclick="javascript:this.src='Captcha.asp?'+Math.random();" style="width:46px;height:14px;border:0px;vertical-align:middle;cursor:pointer;" /><%end if%></td><td style="width:auto;"><span id="JSinnerHTML_AddModify" class="BlackS14"><%=GuSystemMessageContent%></span></td></tr></table></td>
 </tr>
</form>
</table>
<br />
<%end sub%>


<%sub GuestbookList()%>
<form method="post" name="Guestbook_Reply" action="">
<input type="hidden" id="GB_Reply" name="GB_Reply" value="" />
</form>
<table border="0px" cellpadding="0px" cellspacing="1px" class="Style_Table_List_Whole">
<form method="post" name="Guestbook_Notepad_List" action="" onsubmit="javascript:this.action='<%="?Command="&trim(request.querystring("Command"))&"&Type=Guestbook&Status="&trim(request.querystring("Status"))&"&Keyword="&trim(request.querystring("Keyword"))&"&Page="%>'+this.Page.value;">
 <tr>
  <td colspan="2" align="center" class="Style_Table_List_Title"><span class="Style_Title_List">用 户 留 言 列 表</span></td>
 </tr>
 <tr>
  <td colspan="2" class="Style_Table_List_Distance"></td>
 </tr>
<%
	SearchKeywordString=trim(request.querystring("Keyword"))
	SearchKeywordString=GuReplaceSearchKeyword(SearchKeywordString)
	if len(SearchKeywordString)>=1 then
		SQLKeyword="and (GB_UserName like '%"&SearchKeywordString&"%' or GB_UserEMail like '%"&SearchKeywordString&"%' or GB_UserMobilePhone like '%"&SearchKeywordString&"%' or GB_UserIMQQ like '%"&SearchKeywordString&"%' or GB_Content like '%"&SearchKeywordString&"%' or GB_Reply like '%"&SearchKeywordString&"%' or GB_AddIP like '%"&SearchKeywordString&"%' or GB_ModifyIP like '%"&SearchKeywordString&"%')"
		SQLBofEof="<span style=""color:#C0C0C0;"">没有找到相关数据，请重新输入关键字词搜索，<a href=""javascript:void(0);"" onclick=""javascript:window.history.go(-1);"" class=""SilverS14"">［返回］</a></span>"
	else
		SQLKeyword=""
		SQLBofEof="<span style=""color:#C0C0C0;"">没有找到相关数据</span>"
	end if

	if AdministratorID="" or AdministratorAccount="" then
		SQLStatus="and GB_Status='Show'"
	else
		if trim(request.querystring("Status"))="Show" then
			SQLStatus="and GB_Status='Show'"
		elseif trim(request.querystring("Status"))="Hidden" then
			SQLStatus="and GB_Status='Hidden'"
		else
			SQLStatus=""
		end if
	end if

	set ARLR=Server.CreateObject("ADODB.RecordSet")
	SQL="select * from "&GuDatabaseTablePrefix&"_Guestbook where GB_ID>=1 "&SQLKeyword&" "&SQLStatus&" order by GB_ID desc"
	ARLR.open SQL,ACMA,1,1
	if ARLR.bof and ARLR.eof then
		response.write "<tr><td colspan=""2"" align=""left"" class=""Style_Table_List_Content"" style=""height:44px;background:#FFFFFF;"">&nbsp;<span class=""SilverS14"">"&SQLBofEof&"</span></td></tr>"
	else
		PageSize=GuestbookListPageSize
		PageCount=ARLR.recordcount
		PagePresent=trim(request.querystring("Page"))

		if PagePresent="" then
			PagePresent=1
		else
			if isNumeric(PagePresent)=true then
				if PagePresent<=1 then
					PagePresent=1
				else
					PagePresent=fix(PagePresent)
				end if
			else
				PagePresent=1
			end if
		end if

		if (PageCount mod PageSize)=0 then
			PageTotal=PageCount\PageSize
		else
			PageTotal=PageCount\PageSize+1
		end if

		if PageTotal<=1 then
			PageTotal=1
		end if

		if PagePresent>=PageTotal then
			PagePresent=PageTotal
		end if

		PageOrder=1
		PageStart=(PagePresent-1)*PageSize+1

		do while not ARLR.eof
			if PageOrder>=PageStart+PageSize then
				exit do
			end if
			if PageOrder>=PageStart then

				set ARRS=Server.CreateObject("ADODB.RecordSet")
				SQL="select top 1 * from "&GuDatabaseTablePrefix&"_Guestbook where GB_ID<"&ARLR("GB_ID")&" "&SQLKeyword&" "&SQLStatus&" order by GB_ID desc"
				ARRS.open SQL,ACMA,1,1
				if ARRS.bof and ARRS.eof then
					GuestbookIDNext="0"
				else
					GuestbookIDNext=ARRS("GB_ID")
				end if
				ARRS.close
				set ARRS=nothing
%>
 <tr>
  <td align="center" class="Style_Table_List_Name"  style="width:90px;"><span class="BlackS14">用户信息</span></td>
  <td align="left" class="Style_Table_List_Name" style="width:867px;"><table border="0px" cellpadding="0px" cellspacing="0px" style="width:867px;height:auto;"><tr>
	<td align="left" style="width:707px;height:auto;padding:0px 0px 0px 10px;"><span class="BlackS14">昵称：<span style="font-family:Verdana,Arial;"><%=ARLR("GB_UserName")%></span> | 电子邮箱：<%if AdministratorID="" then%><span style="font-family:Verdana,Arial;">****</span><%else%><a href="mailto:<%=ARLR("GB_UserEMail")%>" target="_blank" class="BlackV14"><%=ARLR("GB_UserEMail")%></a><%end if%> | 手机号码：<%if AdministratorID="" then%><span style="font-family:Verdana,Arial;">****</span><%else%><span style="font-family:Verdana,Arial;"><%=ARLR("GB_UserMobilePhone")%></span><%end if%> | <span style="font-family:Verdana,Arial;">QQ</span>账号：<%if AdministratorID="" then%><span style="font-family:Verdana,Arial;">****</span><%else%><span style="font-family:Verdana,Arial;"><%=ARLR("GB_UserIMQQ")%></span><%end if%> | <span style="font-family:Verdana,Arial;">IP</span>：<%if AdministratorID="" then%><span style="font-family:Verdana,Arial;"><%=split(ARLR("GB_AddIP")&"...",".")(0)%>.<%=split(ARLR("GB_AddIP")&"...",".")(1)%>.*.*</span><%else%><a href="<%=GuBrowserIPAddressKuaidial%><%=ARLR("GB_AddIP")%>" target="_blank" class="BlackV14"><%=ARLR("GB_AddIP")%></a><%end if%></span></td>
	<td align="right" style="width:120px;height:auto;"><a href="javascript:void(0);" onclick="javascript:GuElementStyleDisplayBAB1or('Guestbook_Reply_Form_<%=ARLR("GB_ID")%>','Guestbook_Reply_Content_<%=ARLR("GB_ID")%>');Guestbook_Reply_Check('GB_Reply_<%=ARLR("GB_ID")%>');" class="GrayS14">［回复］</a><a href="<%=WebpageName%>?Command=List&Type=Guestbook&Status=<%=ARLR("GB_Status")%>" class="<%if ARLR("GB_Status")="Hidden" then response.write "SilverS14" else response.write "BlackS14" end if%>">［<%=GuResourceStatusText(ARLR("GB_Status"))%>］</a></td>
	<td align="center" style="width:30px;height:auto;padding:3px 4px 0px 0px;"><input type="checkbox" id="ResourceID" name="ResourceID" value="<%=ARLR("GB_ID")%>" class="Style_InputCheckbox" /></td></tr></table></td>
 </tr>
 <tr>
  <td align="center" class="Style_Table_List_Content" style="background:#FFFFFF;vertical-align:top;"><table border="0px" cellpadding="0px" cellspacing="0px" style="width:100%;height:auto;"><tr><td align="center" style="width:auto;height:35px;"><span class="BlackS14">留言内容</span></td></tr><tr><td class="Style_Table_List_Whole" style="width:auto;height:1px;"></td></tr></table></td>
  <td align="left" class="Style_Table_List_Content" style="background:#FFFFFF;vertical-align:top;"><table border="0px" cellpadding="0px" cellspacing="0px" style="width:100%;height:auto;table-layout:fixed;word-break:break-all;">
  <tr>
	<td align="left" style="width:851px;height:auto;padding:7px 0px 4px 10px;"><span class="BlackS14" style="line-height:150%;"><%=Server.HTMLEncode(ARLR("GB_Content"))%>&nbsp;&nbsp;<%=ARLR("GB_AddDateTime")%></span></td>
	<td rowspan="2" align="right" style="width:6px;height:auto;padding:0px 0px 0px 0px;vertical-align:bottom;"><a name="<%=GuestbookIDNext%>" style="font-size:2px;color:#FFFFFF;">.</a></td>
  </tr>
  <tr id="Guestbook_Reply_Content_<%=ARLR("GB_ID")%>" style="display:<%if len(ARLR("GB_Reply"))>=1 then response.write "" else response.write "none" end if%>;">
	<td align="left" style="width:auto;height:2px;padding:0px 0px 8px 10px;vertical-align:bottom;"><span class="NavyS14" style="line-height:140%;"><%=Server.HTMLEncode(ARLR("GB_Reply"))%>&nbsp;&nbsp;<span id="" class="NavyS14" style="display:none;"><%=ARLR("GB_ReplyAdministrator")%><%=ARLR("GB_ReplyIP")%></span><%=ARLR("GB_ReplyDateTime")%></span></td>
  </tr>
  <tr id="Guestbook_Reply_Form_<%=ARLR("GB_ID")%>" style="display:none;">
	<td align="left" style="width:auto;height:26px;padding:0px 0px 8px 10px;"><table border="0px" cellpadding="0px" cellspacing="0px" style="width:auto;height:auto;"><tr><td align="left" style="width:610px;height:auto;"><input type="text" id="GB_Reply_<%=ARLR("GB_ID")%>" name="GB_Reply_<%=ARLR("GB_ID")%>" maxlength="200" value="<%=Server.HTMLEncode(ARLR("GB_Reply"))%>" class="Style_InputText" style="width:600px;" /></td><td align="left" style="width:50px;height:auto;"><input type="button" id="" name="Button_Reply" value="回复" onclick="javascript:Guestbook_Reply_Submit(<%=ARLR("GB_ID")%>);" class="Style_InputText" style="width:40px;height:25px;background:#EBE9ED;font-family:宋体,新宋体;font-size:12px;" /></td></tr></table></td>
  </tr></table></td>
 </tr>
<%
			end if
			PageOrder=PageOrder+1
		ARLR.movenext
		loop
%>
 <tr>
  <td colspan="2" align="right" class="Style_Table_List_Operate"><table border="0px" cellpadding="0px" cellspacing="0px" style="width:948px;height:auto;"><tr><td align="left" style="width:220px;height:auto;"><span id="JSinnerHTML_List" class="BlackS14"><%=GuSystemMessageContent%></span></td><td align="right" style="width:140px;height:auto;"><input type="text" id="Keyword" name="Keyword" maxlength="100" value="<%=trim(request.querystring("Keyword"))%>" onfocus="javascript:this.select();" class="Style_InputText" style="width:132px;" /></td><td align="right" style="width:588px;height:auto;"><input type="button" id="" name="Button_Search" value="" onclick="javascript:Guestbook_Search_Check(this.form,'Guestbook');" class="Style_Button_Search" />&nbsp;<input type="button" id="" name="Button_Reload" value="" onclick="javascript:window.location.reload();" class="Style_Button_Reload" />&nbsp;<input type="button" id="" name="Button_Show" value="" onclick="javascript:GuSubmitConfirm('确定设置显示所有选择的留言吗？',this.form,'?Command=Show&Type=Guestbook&Status=<%=trim(request.querystring("Status"))%>&Keyword=<%=trim(request.querystring("Keyword"))%>&Page=<%=trim(request.querystring("Page"))%>',this.form.ResourceID);" class="Style_Button_Show" />&nbsp;<input type="button" id="" name="Button_Hidden" value="" onclick="javascript:this.form.action='?Command=Hidden&Type=Guestbook&Status=<%=trim(request.querystring("Status"))%>&Keyword=<%=trim(request.querystring("Keyword"))%>&Page=<%=trim(request.querystring("Page"))%>';this.form.submit();" class="Style_Button_Hidden" />&nbsp;<input type="button" id="" name="Button_Delete" value="" onclick="javascript:GuSubmitConfirm('确定删除所有选择的留言吗？',this.form,'?Command=Delete&Type=Guestbook&Status=<%=trim(request.querystring("Status"))%>&Keyword=<%=trim(request.querystring("Keyword"))%>&Page=<%=trim(request.querystring("Page"))%>',this.form.ResourceID);" class="Style_Button_Delete" />&nbsp;<input type="button" id="" name="Button_Select_All" value="" onclick="javascript:GuElementCheckedAll(this.form.ResourceID);" class="Style_Button_Select_All" />&nbsp;<input type="button" id="" name="Button_Select_Reverse" value="" onclick="javascript:GuElementCheckedReverse(this.form.ResourceID);" class="Style_Button_Select_Reverse" />&nbsp;<input type="button" id="" name="Button_Select_Reverse" value="" onclick="javascript:GuElementCheckedClear(this.form.ResourceID);" class="Style_Button_Select_Clear" />&nbsp;</td></tr></table></td>
 </tr>
 <tr>
  <td colspan="2" align="left" class="Style_Table_List_Content" style="background:#FFFFFF;"><%
	response.write "<table border=""0px"" cellpadding=""0px"" cellspacing=""0px"" style=""width:100%;height:auto;"">"
	response.write " <tr>"
	response.write "  <td align=""left"" style=""width:260px;height:auto;padding:10px 0px 6px 8px;""><div id="""" class=""Style_Pagination_Admin_BackNext"">"

	if PagePresent<=1 then
		response.write "<span>第一页</span>&nbsp;<span>上一页</span>&nbsp;"
	else
		response.write "<a href=""?Command=List&Type=Guestbook&Status="&trim(request.querystring("Status"))&"&Keyword="&trim(request.querystring("Keyword"))&"&Page=1"">第一页</a>&nbsp;"
		response.write "<a href=""?Command=List&Type=Guestbook&Status="&trim(request.querystring("Status"))&"&Keyword="&trim(request.querystring("Keyword"))&"&Page="&(PagePresent-1)&""">上一页</a>&nbsp;"
	end if

	if PagePresent>=PageTotal then
		response.write "<span>下一页</span>&nbsp;<span>最末页</span>&nbsp;"
	else
		response.write "<a href=""?Command=List&Type=Guestbook&Status="&trim(request.querystring("Status"))&"&Keyword="&trim(request.querystring("Keyword"))&"&Page="&(PagePresent+1)&""">下一页</a>&nbsp;"
		response.write "<a href=""?Command=List&Type=Guestbook&Status="&trim(request.querystring("Status"))&"&Keyword="&trim(request.querystring("Keyword"))&"&Page="&PageTotal&""">最末页</a>"
	end if

	response.write "</div></td>"
	response.write "  <td align=""right"" style=""width:auto;height:auto;padding:10px 8px 6px 0px;""><div id="""" class=""Style_Pagination_Admin_Number"">"

	if PageTotal<=9 then
		PageStart=1
		PageEnd=PageTotal
	else
		if PagePresent<=5 then
			PageStart=1
			PageEnd=PagePresent+(9-PagePresent)
			if PageEnd>=PageTotal then
				PageEnd=PageTotal
			end if
		else
			if PageTotal-PagePresent<=3 then
				PageStart=PagePresent-(9-(PageTotal-PagePresent+1))
			else
				PageStart=PagePresent-4
			end if
			if PagePresent+4>=PageTotal then
				PageEnd=PageTotal
			else
				PageEnd=PagePresent+4
			end if
		end if
	end if

	for PageNumber=PageStart to PageEnd step 1
	if PageNumber=PagePresent then
		response.write "&nbsp;<a href=""#"" class=""Present"">"&PageNumber&"</a>"
	else
		response.write "&nbsp;<a href=""?Command=List&Type=Guestbook&Status="&trim(request.querystring("Status"))&"&Keyword="&trim(request.querystring("Keyword"))&"&Page="&PageNumber&""">"&PageNumber&"</a>"
	end if
	next

	response.write "</div></td>"
	response.write " </tr>"
	response.write " <tr>"
	response.write "  <td colspan=""2"" align=""right"" style=""width:auto;height:auto;padding:4px 8px 6px 0px;"">"
	response.write "<span class=""BlackS12"">共有<span class=""LightgrayV12"">"&PageCount&"</span>条留言，<span class=""LightgrayV12"">"&PageSize&"</span>条留言/页，</span>"
	response.write "<span class=""BlackS12"">页次：<span class=""LightgrayV12"">"&PagePresent&"<span class=""LightgrayS12"">/</span>"&PageTotal&"</span>，</span>"
	response.write "<span class=""BlackS12"">转到：</span>"
	response.write "<input type=""text"" id=""Page"" name=""Page"" maxlength=""4"" value="""&PagePresent&""" onfocus=""javascript:this.select();"" onblur=""javascript:window.location.href='?Command=List&Type=Guestbook&Status="&trim(request.querystring("Status"))&"&Keyword="&trim(request.querystring("Keyword"))&"&Page='+this.value;"" onafterpaste=""javascript:this.value=this.value.replace(/\D/g,'');"" onkeyup=""javascript:this.value=this.value.replace(/\D/g,'');"" class=""Style_Pagination_Admin_InputText"" />"
	response.write "<input type=""submit"" id="""" name=""Button_Submit"" value="""" class=""Style_Pagination_Admin_Submit"" />"
	response.write "  </td>"
	response.write " </tr>"
	response.write "</table>"
%></td>
 </tr>
<%
	end if
	ARLR.close
	set ARLR=nothing
%>
</form>
</table>
<%end sub%>


<%sub AdministratorModifyForm(AdministratorModifyBoolean)%>
<table border="0px" cellpadding="0px" cellspacing="1px" class="Style_Table_Edit_Whole">
<form method="post" name="Administrator_AddModify" action="?Command=Modify&Type=Account&ID=<%=trim(request.querystring("ID"))%>" onsubmit="javascript:return Administrator_AddModify_Check(this,'Account');">
 <tr>
  <td colspan="4" align="center" class="Style_Table_Edit_Title"><span class="Style_Title_Form">账 号 密 码 设 置</span></td>
 </tr>
 <tr>
  <td colspan="4" class="Style_Table_Edit_Distance"></td>
 </tr>
 <tr>
  <td align="center" class="Style_Table_Edit_Name"><span class="BlackS14">管理账号</span></td>
  <td align="left" class="Style_Table_Edit_Form">&nbsp;<input type="text" id="AM_Account" name="AM_Account" maxlength="40" value="<%=AdministratorAccount%>" onfocus="javascript:this.select();" class="Style_InputText" style="width:367px;background:#EAFBF5;" /></td>
  <td align="center" class="Style_Table_Edit_Name"><span class="BlackS14">验证密码</span></td>
  <td align="left" class="Style_Table_Edit_Form" style="width:auto;">&nbsp;<input type="password" id="AM_Password" name="AM_Password" maxlength="40" value="" onfocus="javascript:this.select();" class="Style_InputText" style="width:367px;background:#EAFBF5;" /></td>
 </tr>
 <tr>
  <td align="center" class="Style_Table_Edit_Name"><span class="BlackS14">新的密码</span></td>
  <td align="left" class="Style_Table_Edit_Form">&nbsp;<input type="password" id="AM_PasswordNew" name="AM_PasswordNew" maxlength="40" value="" onfocus="javascript:this.select();" class="Style_InputText" style="width:367px;" /></td>
  <td align="center" class="Style_Table_Edit_Name"><span class="BlackS14">确认密码</span></td>
  <td align="left" class="Style_Table_Edit_Form" style="width:auto;">&nbsp;<input type="password" id="AM_PasswordConfirm" name="AM_PasswordConfirm" maxlength="40" value="" onfocus="javascript:this.select();" class="Style_InputText" style="width:367px;" /></td>
 </tr>
 <tr>
  <td colspan="4" align="left" class="Style_Table_Edit_Operate"><table border="0px" cellpadding="0px" cellspacing="0px" style="width:100%;height:auto;"><tr><td align="left" style="width:98px;"></td><td align="left" style="width:160px;"><input type="submit" id="" name="Button_Submit" value="" class="Style_Button_Modify" />&nbsp;<input type="reset" id="" name="Button_Reset" value="" class="Style_Button_Reset" /></td><td style="width:auto;">&nbsp;&nbsp;<span id="JSinnerHTML_Account" class="BlackS14"><%if trim(request.querystring("Command"))="Add" or trim(request.querystring("Command"))="Modify" then%><%=GuSystemMessageContent%><%end if%></span></td></tr></table></td>
 </tr>
</form>
</table>
<br />
<%end sub%>


<%sub AdministratorLogList()%>
<table border="0px" cellpadding="0px" cellspacing="1px" class="Style_Table_List_Whole">
<form method="post" name="Admin_Administrator_Log" action="" onsubmit="javascript:this.action='<%="?Command=Edit&Type=Account&Keyword="&trim(request.querystring("Keyword"))&"&Page="%>'+this.Page.value;">
 <tr>
  <td align="center" class="Style_Table_List_Name" style="width:90px;"><span class="BlackS14">编号</span></td>
  <td align="center" class="Style_Table_List_Name" style="width:200px;"><span class="BlackS14">管理账号</span></td>
  <td align="center" class="Style_Table_List_Name" style="width:126px;"><span class="BlackS14">登录<span style="font-family:Verdana;">IP</span>地址</span></td>
  <td align="center" class="Style_Table_List_Name" style="width:152px;"><span class="BlackS14">登录日期时间</span></td>
  <td align="center" class="Style_Table_List_Name" style="width:341px;"><span class="BlackS14">备注</span></td>
  <td align="center" class="Style_Table_List_Name" style="width:44px;"><span class="BlackS14">选择</span></td>
 </tr>
<%
set ARML=Server.CreateObject("ADODB.RecordSet")
SQL="select * from "&GuDatabaseTablePrefix&"_Administrator_Log where AL_ID>=1 order by AL_ID desc"
ARML.open SQL,ACMA,1,1
if ARML.bof and ARML.eof then
	response.write "<tr><td colspan=""6"" align=""left"" style=""width:auto;height:40px;background:#FFFFFF;"">&nbsp;<span class=""SilverS14"">没有找到相关数据</span></td></tr>"
else
	PageSize=AdministratorLogListPageSize
	PageCount=ARML.recordcount
	PagePresent=trim(request.querystring("Page"))

	if PagePresent="" then
		PagePresent=1
	else
		if isNumeric(PagePresent)=true then
			if PagePresent<=1 then
				PagePresent=1
			else
				PagePresent=fix(PagePresent)
			end if
		else
			PagePresent=1
		end if
	end if

	if (PageCount mod PageSize)=0 then
		PageTotal=PageCount\PageSize
	else
		PageTotal=PageCount\PageSize+1
	end if

	if PageTotal<=1 then
		PageTotal=1
	end if

	if PagePresent>=PageTotal then
		PagePresent=PageTotal
	end if

	PageOrder=1
	PageStart=(PagePresent-1)*PageSize+1

	do while not ARML.eof
		if PageOrder>=PageStart+PageSize then
			exit do
		end if
		if PageOrder>=PageStart then
%>
 <tr bgcolor="#FFFFFF" onmouseover="javascript:this.style.backgroundColor='#EAFBF5';" onmouseout="javascript:this.style.backgroundColor='';">
  <td align="center" class="Style_Table_List_Content"><span class="BlackS14"><%=ARML("AL_ID")%></span></td>
  <td align="left" class="Style_Table_List_Content">&nbsp;<span class="BlackV14"><%=ARML("AL_AddAdministrator")%></span></td>
  <td align="center" class="Style_Table_List_Content"><a href="<%=GuBrowserIPAddressKuaidial%><%=ARML("AL_AddIP")%>" target="_blank" class="BlackS14"><%=ARML("AL_AddIP")%></a></td>
  <td align="center" class="Style_Table_List_Content"><span class="BlackS14"><%=ARML("AL_AddDateTime")%></span></td>
  <td align="left" class="Style_Table_List_Content">&nbsp;<span class="BlackS14"><%=ARML("AL_Note")%></span></td>
  <td align="center" class="Style_Table_List_Content"><input type="checkbox" id="GuResourceID" name="GuResourceID" value="<%=ARML("AL_ID")%>" class="Style_InputCheckbox" /></td>
 </tr>
<%
		end if
		PageOrder=PageOrder+1
	ARML.movenext
	loop
%>
 <tr>
  <td colspan="6" align="right" class="Style_Table_List_Operate"><table border="0px" cellpadding="0px" cellspacing="0px" style="width:948px;height:auto;"><tr><td align="left" style="width:648px;height:auto;"><span id="JSinnerHTML_List" class="BlackS14"><%if trim(request.querystring("Command"))="LogDelete" then%><%=GuSystemMessageContent%><%end if%></span></td><td align="right" style="width:300px;height:auto;"><input type="button" id="" name="Button_Delete" value="" onclick="javascript:GuSubmitConfirm('确定删除所有选择的管理账号日志吗？',this.form,'?Command=LogDelete&Type=Account&Page=<%=trim(request.querystring("Page"))%>',this.form.GuResourceID);" class="Style_Button_Delete" />&nbsp;<input type="button" id="" name="Button_Select_All" value="" onclick="javascript:GuElementCheckedAll(this.form.GuResourceID);" class="Style_Button_Select_All" />&nbsp;<input type="button" id="" name="Button_Select_Reverse" value="" onclick="javascript:GuElementCheckedReverse(this.form.GuResourceID);" class="Style_Button_Select_Reverse" />&nbsp;<input type="button" id="" name="Button_Clear" value="" onclick="javascript:GuElementCheckedClear(this.form.GuResourceID);" class="Style_Button_Select_Clear" />&nbsp;</td></tr></table></td>
 </tr>
 <tr>
  <td colspan="6" align="left" class="Style_Table_List_Content" style="background:#FFFFFF;"><%
	response.write "<table border=""0px"" cellpadding=""0px"" cellspacing=""0px"" style=""width:100%;height:auto;"">"
	response.write " <tr>"
	response.write "  <td align=""left"" style=""width:260px;height:auto;padding:10px 0px 6px 8px;""><div id="""" class=""Style_Pagination_Admin_BackNext"">"

	if PagePresent<=1 then
		response.write "<span>第一页</span>&nbsp;<span>上一页</span>&nbsp;"
	else
		response.write "<a href=""?Command=Edit&Type=Account&Keyword="&trim(request.querystring("Keyword"))&"&Page=1"">第一页</a>&nbsp;"
		response.write "<a href=""?Command=Edit&Type=Account&Keyword="&trim(request.querystring("Keyword"))&"&Page="&(PagePresent-1)&""">上一页</a>&nbsp;"
	end if

	if PagePresent>=PageTotal then
		response.write "<span>下一页</span>&nbsp;<span>最末页</span>&nbsp;"
	else
		response.write "<a href=""?Command=Edit&Type=Account&Keyword="&trim(request.querystring("Keyword"))&"&Page="&(PagePresent+1)&""">下一页</a>&nbsp;"
		response.write "<a href=""?Command=Edit&Type=Account&Keyword="&trim(request.querystring("Keyword"))&"&Page="&PageTotal&""">最末页</a>"
	end if

	response.write "</div></td>"
	response.write "  <td align=""right"" style=""width:auto;height:auto;padding:10px 8px 6px 0px;""><div id="""" class=""Style_Pagination_Admin_Number"">"

	if PageTotal<=9 then
		PageStart=1
		PageEnd=PageTotal
	else
		if PagePresent<=5 then
			PageStart=1
			PageEnd=PagePresent+(9-PagePresent)
			if PageEnd>=PageTotal then
				PageEnd=PageTotal
			end if
		else
			if PageTotal-PagePresent<=3 then
				PageStart=PagePresent-(9-(PageTotal-PagePresent+1))
			else
				PageStart=PagePresent-4
			end if
			if PagePresent+4>=PageTotal then
				PageEnd=PageTotal
			else
				PageEnd=PagePresent+4
			end if
		end if
	end if

	for PageNumber=PageStart to PageEnd step 1
	if PageNumber=PagePresent then
		response.write "&nbsp;<a href=""#"" class=""Present"">"&PageNumber&"</a>"
	else
		response.write "&nbsp;<a href=""?Command=Edit&Type=Account&Keyword="&trim(request.querystring("Keyword"))&"&Page="&PageNumber&""">"&PageNumber&"</a>"
	end if
	next

	response.write "</div></td>"
	response.write " </tr>"
	response.write " <tr>"
	response.write "  <td colspan=""2"" align=""right"" style=""width:auto;height:auto;padding:4px 8px 6px 0px;"">"
	response.write "<span class=""BlackS12"">共有<span class=""LightgrayV12"">"&PageCount&"</span>条日志，<span class=""LightgrayV12"">"&PageSize&"</span>条日志/页，</span>"
	response.write "<span class=""BlackS12"">页次：<span class=""LightgrayV12"">"&PagePresent&"<span class=""LightgrayS12"">/</span>"&PageTotal&"</span>，</span>"
	response.write "<span class=""BlackS12"">转到：</span>"
	response.write "<input type=""text"" id=""Page"" name=""Page"" maxlength=""4"" value="""&PagePresent&""" onfocus=""javascript:this.select();"" onblur=""javascript:window.location.href='?Command=Edit&Type=Account&Keyword="&trim(request.querystring("Keyword"))&"&Page='+this.value;"" onafterpaste=""javascript:this.value=this.value.replace(/\D/g,'');"" onkeyup=""javascript:this.value=this.value.replace(/\D/g,'');"" class=""Style_Pagination_Admin_InputText"" />"
	response.write "<input type=""submit"" id="""" name=""Button_Submit"" value=""Submit"" class=""Style_Pagination_Admin_Submit"" />"
	response.write "  </td>"
	response.write " </tr>"
	response.write "</table>"
%></td>
 </tr>
<%
end if
ARML.close
set ARML=nothing
%>
</form>
</table>
<%end sub%>


<%sub GuestbookJSAName(ANameValue)%>
<script type="text/javascript">
<!--
	setTimeout("document.location.href='#<%=ANameValue%>'",0);
//-->
</script>
<%end sub%>


<br />
<a name="GuestbookBottom" style="font-size:2px;color:#FFFFFF;">.</a>
<br />
<br />

<%
ACMA.close
set ACMA=nothing
%>

<!--  啊估留言簿 V2.1 20190114  -->

</div>
</body>
</html>

