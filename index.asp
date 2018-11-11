<!--#include virtual="/config.asp" -->
<!--#include virtual="/inc_conn_open.asp" -->
<!--#include virtual="/inc_get_data.asp" -->
<!--#include virtual="/new/inc_header.asp" -->
<!--#include virtual="/new/inc_menu_top.asp" -->

<%
' ################################################
' FRED FRED FRED FRED FRED FRED FRED FRED FRED FRED 
' // KILL ANY EXISTING TRACKERID - NEW TRANSACTION
' ################################################
Session.Contents.Remove("ab_TrackerID")
Session.Contents.Remove("ab_AFSFree")

If Request.QueryString("total") = 1 Then
	' GET REVENUE AND TOTALS FOR TODAY
	' ################################
	Set RS = Server.CreateObject("ADODB.Recordset")
	QueryRevenue = "SELECT SUM(tblAdvertsMain.TotalCost) AS TotalRevenue FROM tblAdvertsMain WHERE tblAdvertsMain.Status = 'OK' AND tblAdvertsMain.DateVisit = '" & FormatDateTime(Date,1) & "' "
	RS.Open QueryRevenue, Connect, adOpenStatic, adLockOptimistic
	TotalRevenueToday = RS("TotalRevenue")
	RS.Close
End If

' ##############################
' SHOW PLACEHOLDER IF ONE IS SET
' ##############################
If PlaceHolder = True Then
	Response.Write PlaceHolderText
Else %>

    <div class="panel panel-default">
        <div class="panel-body">
            <h2>Place an advert in <%= TitleHead %> and online</h2>
            <p>Select the category of advert you would like to place and follow the instructions.<br />
            You can pay by Mastercard, Visa, Maestro and Electon credit and debit cards. <%= TotalRevenueToday %></p>
        </div>
    </div>

    <div class="row">

    <%
    arrThisPubCategories = Split(Replace(ThisPubCategories," ",""),",")
    For a = 0 To UBound(arrThisPubCategories)

        ' ####################################################################################
        ' // GET THE CATEGORIES USED IN THIS PUBLICATION. USE ANY CUSTOM CATEGORY ORDER
        ' ####################################################################################
        Set RS = Server.CreateObject("ADODB.Recordset")
        Query = "SELECT ID, CategoryName, CategoryImage, Blurb, PricingInfo, LinkText, LinkFolder, CFGroupID, LinkOnly " &_
                "FROM tblCategoryLibrary " &_
                "WHERE ID = " & arrThisPubCategories(a)
        RS.Open Query, Connect, adOpenStatic, adLockOptimistic
        CatID = RS("ID")
        CategoryName = RS("CategoryName")
        CategoryImage = RS("CategoryImage")
        'Blurb = Left(RS("Blurb"),Session("ab_max_blurb"))
		'PricingInfo = Left(RS("PricingInfo"),Session("ab_max_pricinginfo"))
		Blurb = RS("Blurb")
        PricingInfo = RS("PricingInfo")
        LinkText = RS("LinkText")
        LinkFolder = RS("LinkFolder")
        CFGroupID = RS("CFGroupID")
        LinkOnly = RS("LinkOnly")

        ' ####################################################################################
        ' // OVERWRITE ANY CUSTOM TEXT FOR THESE CATEGORIES THAT HAS BEEN SET BY A CENTRE
        ' ####################################################################################
        Set RS_PubText = Server.CreateObject("ADODB.Recordset")
        Query_PubText = "SELECT CategoryName, CategoryImage, Blurb, PricingInfo, LinkText FROM tblPublicationsText WHERE PubID = " & ThisPublicationID & " AND CatID = " & CatID
        RS_PubText.Open Query_PubText, Connect, adOpenStatic, adLockOptimistic
        If Not RS_PubText.EOF Then
            CategoryName = RS_PubText("CategoryName")
            Blurb = Left(RS_PubText("Blurb"),max_blurb)
            PricingInfo = Left(RS_PubText("PricingInfo"),max_pricinginfo)
            LinkText = RS_PubText("LinkText")
            If Len(RS_PubText("PricingInfo")) > max_pricinginfo Or Len(RS_PubText("Blurb")) > max_blurb Then button_state = "danger" Else button_state = "success" End If
        End If
        RS_PubText.Close
        Set RS_PubText = Nothing
		
        ' ####################################################################################
        ' // OVERWRITE IF THIS IS A RECRUITEMENT CATEGORY AND INSTEAD LINK OUT TO X1jobs etc.
        ' ####################################################################################
		If Request.QueryString("newjobs") = "true" Then
			If CategoryName = "Recruitment" Then
				Blurb = "With over 2 million visits across our network every day, you're bound to find your ideal candidate in no time."
				PricingInfo = "Prices start from £99 + VAT"
				LinkFolder = "https://www." & JobSite & ".com/recruiters/post-a-job"
			End If
		End If
    %>






        <!-- MAIN CATEGORY PANEL -->
        <div class="col-sm-6 col-md-4 col-lg-4">
            <div class="panel panel-default">
                <div class="panel-heading panel-custom text-center">
                    <%= CategoryName %>
                </div>

                <a href="<%= LinkFolder %>/">
                    <div class="panel-body text-center panel-height-category">
                        <img src="/catimages/<%= CategoryImage %>" alt="<%= LinkText %>" width="310" height="150" border="0" class="img-responsive center-block" /><br />
                        <strong><%= Blurb %></strong><br />
                        <%= PricingInfo %>
                    </div>
                </a>

                    <div class="panel-footer text-center">
                        <a href="<%= LinkFolder %>/" class="btn btn-lg btn-primary" role="button">Select this category <span class="glyphicon glyphicon-chevron-right"></span></a>
                        <%
                        If Request.Cookies("ab_ShowEdits") = "true" Then %>
                            <a target="_blank" href="http://adbookeradmin.newsquest.co.uk/admin_pubtext.asp?pubID=<%= ThisPublicationID %>&catID=<%= CatID %>" class="btn btn-<%= button_state %> btn-xs" role="button"><span class="glyphicon glyphicon-pencil"></span></a>
                        <% End If %>
                    </div>
            </div>
        </div>
        <!-- /MAIN CATEGORY PANEL -->




    <%
    Next
    RS.Close
    Set RS = Nothing
    %>

    </div>

    <%
    ' ####################################################
    ' // LOOK FOR ANY LINK CATEGORIES FOR THIS PUBLICATION
    ' ####################################################
    Set RS = Server.CreateObject("ADODB.Recordset")
    Query = "SELECT * FROM tblCategoryOthers WHERE PubID = " & ThisPublicationID & " AND DisplayOrder IS NOT NULL ORDER BY DisplayOrder"
    RS.Open Query, Connect, adOpenStatic, adLockOptimistic
    %>

    <!-- LINK CATEGORIES -->
    <%
    If ThisPublication = "newsshopper.co.uk" Then
        AdvertisingLink = "http://www.newsshopper.co.uk/advertising/getconnected/"
    Else
        AdvertisingLink = "http://www." & ThisPublication & "/digital_advertising/"
    End If
    %>
    <div class="panel panel-default">
        <div class="panel-heading panel-custom">
            Online advertising
        </div>
        <div class="panel-body">
            Looking for new customers? <a href="<%= AdvertisingLink %>">Advertise online and get your message in front of new potential customers</a>
        </div>
    </div>

    <% Do Until RS.EOF %>
    <div class="panel panel-default">
        <div class="panel-heading panel-custom">
            <%= RS("CategoryName") %>
        </div>
        <div class="panel-body">
            <% If RS("Blurb") <> "" Then Response.Write RS("Blurb") & "<br />" End If %>
            <% If RS("PricingInfo") <> "" Then Response.Write RS("PricingInfo") & "<br />" End If %>
            <a href="<%= RS("LinkToURL") %>"><%= RS("LinkText") %> &raquo;</a>
        </div>
    </div>
    <% RS.MoveNext
       Loop
    %>
    <!-- /LINK CATEGORIES -->

<% End If %>

<!--#include virtual="/phone_data.asp" -->
<!--#include virtual="/inc_conn_close.asp" -->
<!--#include virtual="/new/inc_footer.asp" -->
