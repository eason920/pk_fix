
<div id="fb-root"></div>
<style>
  /* body */
  .breadcrumbs + .nk-box {
    background-image: url(./assets/images/four_color_icon.png);
    background-position: top right;
    background-repeat: no-repeat;
  }

  /* all */
  .nk-navbar .dropdown a{
    border-bottom: solid 2px #f1a930;
  }

  /* mb */
  .nk-navbar-side .dropdown li{
    width: 80%;
    margin: 0 auto;
  }
  
  .nk-navbar-side .dropdown li:first-child a{
    border: none;
  }

  @media screen and (max-width: 767px){
    /* body */
    .breadcrumbs + .nk-box {
      background-size: 50%;
    }
  }
</style>
<script>(function(d, s, id) {
  var js, fjs = d.getElementsByTagName(s)[0];
  if (d.getElementById(id)) return;
  js = d.createElement(s); js.id = id;
  js.src = "//connect.facebook.net/zh_TW/sdk.js#xfbml=1&version=v2.10&appId=<%= a_facebookID %>";
  fjs.parentNode.insertBefore(js, fjs);
}(document, 'script', 'facebook-jssdk'));</script>


<% Dim appLAmenu, appLAmenuItem
If pdtCtrl = 1 Then
  Dim hdrst, sqlhdrst, hd_type_id1
  'productTypes - 取得商品大類
  If pdcCtrl = 1 And pdkCtrl = 1 Then	'兩層分類
	  sqlhdrst = "SELECT * FROM productTypes WHERE openon = 'yes' ORDER BY gp ASC "
  Else								'一層分類
	  sqlhdrst = "SELECT TOP 1 * FROM productTypes WHERE isde = 'no' AND openon = 'yes' ORDER BY gp ASC "
  End If
  Set hdrst = GetMdb("pkgo/PKthinkdb.mdb",sqlhdrst)
  If Not hdrst.EOF Then hd_type_id1 = hdrst("type_id") Else hd_type_id1 = 0 '取得第一筆商品大類編號
End If


If abtCtrl = 1 And inmCtrl = 1 Then	'多頁公司簡介
	Dim hdinrs, hdinrs4, hdinrs5
	'intros - 考試課程
	Set hdinrs = GetMdb("pkgo/PKthinkdb.mdb","SELECT in_id, in_title, in_title2, in_intro FROM intros WHERE type_id = 3 AND openon = 'yes' ORDER BY gp ASC ")
	hdinrsCount = hdinrs.RecordCount

	'intros - 專業課程
	Set hdinrs4 = GetMdb("pkgo/PKthinkdb.mdb","SELECT in_id, in_title, in_title2, in_intro FROM intros WHERE type_id = 4 AND openon = 'yes' ORDER BY gp ASC ")
	hdinrs4Count = hdinrs4.RecordCount

	'intros - 大學線上學分
	Set hdinrs5 = GetMdb("pkgo/PKthinkdb.mdb","SELECT in_id, in_title, in_title2, in_intro FROM intros WHERE type_id = 5 AND openon = 'yes' ORDER BY gp ASC ")
	hdinrs5Count = hdinrs5.RecordCount

	'intros - 美國大學申請
	Set hdinrs6 = GetMdb("pkgo/PKthinkdb.mdb","SELECT in_id, in_title, in_title2, in_intro FROM intros WHERE type_id = 6 AND openon = 'yes' ORDER BY gp ASC ")
	hdinrs6Count = hdinrs6.RecordCount
End If


If newCtrl = 1 And nwtCtrl = 1 Then	'最新消息分類
	Dim hdnrst
	'newsTypes
	Set hdnrst = GetMdb("pkgo/PKthinkdb.mdb","SELECT type_id, type_name, type_pic FROM newsTypes WHERE cate_id = 6 AND openon = 'yes' ORDER BY gp ASC ")
	hdnrstCount = hdnrst.RecordCount
End If


If faqCtrl = 1 And facCtrl = 1 Then	'客戶問答分類
	Dim hdqarst
	'qaTypes
	Set hdqarst = GetMdb("pkgo/PKthinkdb.mdb","SELECT type_id, type_name, type_pic FROM qaTypes WHERE openon = 'yes' ORDER BY gp ASC ")
End If


If abmCtrl = 1 And amtCtrl = 1 Then	'相簿分類
	Dim hdabrsk
	'albumKinds
	Set hdabrsk = GetMdb("pkgo/PKthinkdb.mdb","SELECT kind_id, kind_name, kind_pic FROM albumKinds WHERE type_id = 266 AND openon = 'yes' ORDER BY gp ASC ")
End If


'albumKinds - 關於我們
Set leabrs = GetMdb("pkgo/PKthinkdb.mdb","SELECT a.*, k.kind_name FROM album a LEFT JOIN albumKinds k ON a.kind_id = k.kind_id WHERE a.type_id = 266 AND a.kind_id = 702 AND (a.openon = 'yes' AND k.openon = 'yes') ORDER BY a.gp ASC ")
leabrsCount = leabrs.RecordCount


'filesKinds - 美國大學入學申請
Set lefirsk = GetMdb("pkgo/PKthinkdb.mdb","SELECT kind_id, type_id, kind_name FROM filesKinds WHERE openon = 'yes' ORDER BY gp ASC ")
lefirskCount = lefirsk.RecordCount


Dim inmqrs, inmarq_url, inmqrs_count
If blockHeader = "1" Then		'Header No.1 ----------------------------------------------------------- >> %>
  <header id="Header01" class="nk-header <% If PathNow = "default.asp" Or PathNow = "index.asp" Then %>homepage<% End If %>">
    <!-- #include file="_topTyper.asp" -->
    <nav class="nk-navbar nk-navbar-top nk-navbar-sticky nk-navbar-transparent nk-navbar-autohide">
      <div class="container">
        <div class="align-center PCOnly">
          <div class="nk-gap-1-15"></div>
          <a href="./" class="nk-nav-logo " title="<%= aa_web %>"><img src="images/header/<%= logoPic %>" alt="<%= aa_web %>" title="<%= aa_web %>" /></a>
          <div class="nk-gap-1-15"></div>
        </div>
        <div class="nk-nav-table"> 
          <!-- #include file="_menu01.asp" -->
        </div>
      </div>
    </nav>
  </header>
  
<% ElseIf blockHeader = "2" Then	'Header No.2 ----------------------------------------------------------- >> %>

  <header id="Header02" class="nk-header <% If PathNow = "default.asp" Or PathNow = "index.asp" Then %>homepage<% End If %>">
    <!-- #include file="_topTyper.asp" -->
    <div class="nk-contacts-top">
      <div class="container">
		<% If socCtrl = 1 Then	'社群網站連結 %>
          <div class="nk-contacts-right">
            <ul>
              <% If linCtrl = 1 And a_line <> "" Then			'LINE %>
                <li> <a class="nk-contact-icon" href="https://line.me/R/ti/p/%40<%=a_line %>" title="LINE ID: @<%=a_line %>"> <img src="assets/images/if_line.png" alt="LINE ID: @<%=a_line %>" title="LINE ID: @<%=a_line %>" />&nbsp; ID: <%=a_line %> </a> </li>
              <% End If %>
			  <% If wchCtrl = 1 And a_wechat <> "" Then		'WeChat %>
                <li> <a class="nk-contact-icon" href="wechat-qrcode.asp" target="_blank" title="WeChat ID: <%= a_wechat %>"> <img src="assets/images/if_wechat.png" alt="WeChat" title="WeChat" />&nbsp; ID: <%= a_wechat %> </a>
              <% End If %>
			  <% If fbkCtrl = 1 And a_facebook <> "" Then		'Facebook %>
                <li> <a class="nk-contact-icon" href="<%= a_facebook %>" target="_blank" title="Facebook"> <img src="assets/images/if_facebook.png" alt=""> </a> </li>
              <% End If %>
              <% If ggpCtrl = 1 And a_google <> "" Then		'Google+ %>
                <li> <a class="nk-contact-icon" href="<%= a_google %>" target="_blank" title="Google+"> <span class="ion-social-google"></span> </a> </li>
              <% End If %>
              <% If twtCtrl = 1 And a_twitter <> "" Then		'Twitter %>
                <li> <a class="nk-contact-icon" href="<%= a_twitter %>" target="_blank" title="Twitter"> <span class="ion-social-twitter"></span> </a> </li>
              <% End If %>
              <% If igmCtrl = 1 And a_instagram <> "" Then	'Instagram %>
                <li> <a class="nk-contact-icon" href="<%= a_instagram %>" target="_blank" title="Instagram"> <span class="ion-social-instagram-outline"></span> </a> </li>
              <% End If %>

              <li> <a class="nk-contact-icon" href="http://www.ctenglish.com.tw/EN/" title="EN"> | EN </a> </li>
              <li> <a class="nk-contact-icon" href="http://www.ctenglish.com.tw/" title="繁中"> | 繁中 </a> </li>
              <li> <a class="nk-contact-icon" href="http://www.ctenglish.com.tw/CN/" title="简中"> | 简中 </a> </li>
            </ul>
          </div>
        <% End If %>
      </div>
    </div>
    <nav class="nk-navbar nk-navbar-top nk-navbar-sticky nk-navbar-transparent nk-navbar-autohide">
      <div class="container">
        <div class="nk-nav-table">
          <!-- #include file="_menu02.asp" -->
        </div>
      </div>
    </nav>
  </header>
  
<% ElseIf blockHeader = "3" Then	'Header No.3 ----------------------------------------------------------- >> %>

  <div id="Header03">
    <!-- #include file="_topTyper.asp" -->
    <ul class="toplogo">
      <li> <a href="./" class="nk-nav-logo" title="<%= aa_web %>"> <img src="images/header/<%= logoPic %>" alt="<%= aa_web %>" title="<%= aa_web %>" class="logo" /> </a> </li>
    </ul>
    <ul class="nk-nav-toggler-right">
      <li class="single-icon"> <a href="#" class="nk-navbar-full-toggle" title="MENU"> <span class="nk-icon-burger"> <span class="nk-t-1"></span> <span class="nk-t-2"></span> <span class="nk-t-3"></span> </span> </a> </li>
    </ul>
  </div>
  <nav class="nk-navbar nk-navbar-full nk-navbar-align-center" id="nk-full">
    <div class="nk-nav-table">
      <div class="nk-nav-row-full nk-nav-row">
        <div class="nano">
          <div class="nano-content">
            <div class="nk-nav-table">
			  <!-- #include file="_menu03.asp" -->
            </div>
          </div>
        </div>
      </div>
      <div class="nk-nav-row">
        <div class="nk-nav-social">
          <div class="container">
            <div class="row">
              <div class="col-sm-6 text-sm-left"> <a href="./" title="<%= aa_web %>"> <img src="images/header/<%= logoPic %>" alt="<%= aa_web %>" title="<%= aa_web %>" /></a> </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  </nav>
  
<% End If %>


<% If adrCtrl = 1 Then	'右欄收納廣告 %>
<!-- #include file="_rightSide.asp" -->
<% End If %>

<!-- #include file="_leftSide.asp" --> 
