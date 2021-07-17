<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file="_ph.asp" -->
<!-- #include file="_page.asp" -->
<!-- #include file="_olu.asp" -->
<!-- #include file="_login.asp" -->
<%
login() 'before login location


'Receive paras
Dim type_id, v_year
If Request("type") <> "" Then type_id = CInt(Trim(Replace(Replace(Replace(Request("type"), "-", ""), " ", ""), "'", ""))) Else type_id = 0
If Request("year") <> "" Then v_year = CInt(Trim(Replace(Replace(Replace(Request("year"), "-", ""), " ", ""), "'", ""))) Else v_year = 0
'Response.Write "type_id= " & type_id & "<br />"
'Response.Write "v_year= " & v_year & "<br />"


'定義相關參數
Dim blockTitle, list_url, info_url
blockTitle = "影音專區"
page_title = blockTitle
page_title_eng = "Video"
list_url = "video-list.asp"
info_url = "video-info.asp"


'video
Dim vrs, sql, vrsCount
sql = "SELECT v.*, t.type_name FROM video v LEFT JOIN videoTypes t ON v.type_id = t.type_id WHERE v.type_id <> 7 AND (v.openon = 'yes' AND t.openon = 'yes') AND (v.v_time1 <= #"& nowDate &"# AND v.v_time2 >= #"& nowDate &"#) "
If type_id <> 0 Then
    sql = sql & "AND (v.type_id = "& type_id &") "
End If
If v_year <> 0 Then
    sql = sql & "AND (YEAR(v.v_time) = "& v_year &") "
End If
sql = sql & "ORDER BY v.gp ASC, v.v_time DESC "
'Response.Write sql & "<br />"
'Session("pageSqlStrV") = sql
Set vrs = GetMdb(mainDBPath,sql)
vrsCount = vrs.RecordCount
pg = pages(vrs, 12)


'videoTypes
Dim vrst, sqlt, type_name0, type_pic0
If type_id <> 0 Then
    sqlt = "SELECT type_name, type_pic FROM videoTypes WHERE openon = 'yes' AND type_id = "& type_id &" "
  Set vrst = GetMdb(mainDBPath,sqlt)
  If Not vrst.EOF Then
    type_name0 = vrst("type_name")
    If vrst("type_pic") <> "" Then type_pic0 = vrst("type_pic") Else type_pic0 = "NoPicture_type.jpg"
  Else
    Response.Redirect list_url & "?errMsg=" & Server.URLENCODE("No data")
  End If
  vrst.Close
  Set vrst = Nothing
Else
    type_name0 = blockTitle
    type_pic0 = "NoPicture_type.jpg"
End If


page_title = blockTitle
If type_id <> 0 Then
    page_title = type_name0
End If
If v_year <> 0 Then
    page_title = v_year & " " & page_title
End If


'Update Location & Action
LocationNow = blockTitle
ActionNow = "查看" & page_title
olu = oluObj(LocationNow, ActionNow)


%>
<!DOCTYPE html>
<html lang="zh-Hant-TW">
<head>
<title><%= page_title %> - <%= aa_web %></title>
<meta name="description" content="<%= a_description %>">
<!-- #include file="_meta.asp" -->

</head>
<body>
<div class="animsition">
  <!-- #include file="_header.asp" -->

  <div class="nk-main">

    <% If bagCtrl = 1 Then	'內頁橫幅
	  Dim bannerPageTitle, bannerPageTitle2
	  bannerPageTitle = page_title
	  bannerPageTitle2 = "" %>
      <!-- #include file="_bannerPages2.asp" -->
    <% End If %>

    <!-- 網站導覽列 breadcrumbs ------------------------------------ Start. -->
    <div class="breadcrumbs">
      <div class="container">
        <a href="./" title="首頁"><i class="icon ion-ios-home"></i> 首頁</a>
        <span>/</span> <%= blockTitle %>
      </div>
    </div>
    <!-- 網站導覽列 breadcrumbs -------------------------------------- End. -->

    <div class="nk-box">
      <div class="container">
        <div class="row">
          <div class="col-lg-12 lepro">
            <div class="nk-gap-3"></div>
            <div class="alltitle text-black">
              <div class="subtitle"><%= blockTitle %></div>
              <h2>audio</h2>
              <hr class="titleline" />
            </div>
            <!-- #include file="video-list_grid1.asp" -->
            <div class="nk-gap-2"></div>
            <div class="nk-pagination nk-pagination-center"><%= pg %></div>
            <div class="nk-gap-2"></div>
          </div>
        </div>
      </div>
    </div>
  
    <!-- #include file="_footer.asp" -->
  </div>
  
  <!-- #include file="_sidebtn.asp" -->
  <!-- #include file="_search.asp" -->
</div>

<!-- #include file="_js.asp" -->
</body>
</html>
<!-- #include file="_backpost.asp" -->
