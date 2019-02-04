<cfinclude template="vardefinition.cfm"/>
<cfparam name="NewStatus" default="false">
<CFSET variables.audit = createObject("component","#Request.Library.CFC.DotPath#.apps.ats.audit").init(ODBC="#ODBC#", AuditHome="#AuditHome#", NewStatus="#NewStatus#",businessID=request.businessID)>
<cfset structAppend(variables, variables.audit.ConfigLabels(variables),true)/>
<!---
Filename: audaction.cfm

Author			Date		Action
------------------------------------------------------------------------------------------------
Brien Hodges	08/16/02	App Enhancement #2203 - Allow audit to be Verified
Brien Hodges	08/29/02	Add ability to specify days to be notified when >60
Brien Hodges	11/04/02	Added Blue Permnissions block
Brien Hodges	11/06/02	Alterations to email
Brien Hodges	04/03/03	Add VerifyByDate
Brien Hodges	05/28/03	Send an email when Closure Verification Date is changed
Brien Hodges	07/14/03	PDA Accessable page
Brien Hodges	07/24/03	Send email when verification date is removed
Brien Hodges	10/20/03	FORM.ResponPerson was't submitted. Added check for missing variable
Brien Hodges	10/27/03	Only update VerifyPerson when it has the ability to be changed
Chuck Jody		03/11/04	Added URL.PCView so link back goes to correct place in PPC mode
Chuck Jody		03/11/04	Removed sImgWidth from image...it was squishing it an odd way.
Brien Hodges	03/17/04	Added check for FORM.STATUS because certain encodings can destroy a HTML
Brien Hodges	03/25/04	Strip out the a space to the end of textareas to fix SHIFT-JS encoding that
							 combines the "<" with the last character that way typed in
David Zavalza	03/29/04	New attachment model (siteID tree) implemented
Brien Hodges	04/08/04	Add ability to add closure comment by not actually change the status
Brien Hodges	04/20/04	Addition of SubCOE and Multi CC
Brien Hodges	05/03/04	Allow for each Audit Type to have its own # of days into the future that it
							 can be closed
Chuck Jody		05/26/04	Added <CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#PDA/PDAtopBar.cfm"> for Header to return to IP in PocketPC Mode and made
							 the HTML changes for it to fit.
Brien Hodges	06/17/04	App Enhancement #7082 - CAPA Addition
Brien Hodges	09/20/04	Catch failed inserts and inform the user they should try resubmitting when
							 foreign characters exist and dispatch an email
Brien Hodges	10/05/04	Auto Add Permissions for CV
Angel Cisneros  05/22/06	Copy selected attachments to another site
Brien Hodges	04/11/07	Add New Root Cause options
--->
<!--- moved sImgWidth from image...it was squishing it an  --->
<cfset request.ATSAuditName = ""/><!--- Don't remove --->
<cfparam name="variables.auditNameNumberLabel" default="Action Name/Number"/>
<cfparam name="variables.phoneLabel" default="Contact Phone"/>

<cfset variables.atsAdditionalfields = go.getSetupVar(var='atsAdditionalfields',default="")/>
<cfif isStruct(variables.atsAdditionalfields)>
        <cfset additionalATSFields = variables.atsAdditionalfields/>
    </cfif>
<cfif isJson(variables.atsAdditionalfields)>
	<cfset variables.atsAdditionalfields = deserializeJSON(variables.atsAdditionalfields)/>
	<cfif isStruct(variables.atsAdditionalfields)>
		<cfset additionalATSFields = variables.atsAdditionalfields/>
	</cfif>
</cfif>

<cfset variables.showAuditNameNumber  = go.getSetupVar(var='showAuditNameNumber',default=true)/>
<cfset variables.showContactPhone     = go.getSetupVar(var='showContactPhone',default=true)/>
<cfset variables.citationMaxLength	  = go.getSetupVar(var='citationMaxLength',default='250')/>
<cfset variables.factoryIDHistory 	  = go.getSetupVar(var='factoryIDHistory',default='0',appid=3)/>
<cfset variables.globalATSName = request.atsAuditName>

<cfif showAuditNameNumber eq false>
	<cfset variables.auditNameNumberLabel = "" />
</cfif>
<cfif showContactPhone eq false>
	<cfset variables.phoneLabel = "" />
</cfif>
<!--- set variable to show CoResponPerson in GE Non-Industrial --->
<cfset variables.showCoResponPerson = go.getSetupVar(var="ats_showCoResponPerson",default=false)/>
<cfset variables.isCoResponPersonRequired = go.getSetupVar(var="ats_isCoResponPersonRequired",default=false)/>

<!--- All new params please inside the cfscript, thanks! --->
<cfscript>
	 param name="RootCause1Label"     default="Basic Cause";
     param name="RootCause2Label"     default="Near Root Cause";
     param name="bRootCauseDropDowns" default="false";
</cfscript>

<CFPARAM NAME="CI_LOCK_ATS_FindingType" DEFAULT="">
<CFPARAM NAME="CIHome" DEFAULT="">
<CFPARAM NAME="FORM.ACTION" DEFAULT="">
<CFPARAM NAME="FORM.optional_2" DEFAULT="">
<cfparam name="ByPassExport" default="false">
<cfparam name="FORM.EXPCIMandatory" default="">

<cfparam name="findingNotificationRequiredByType" default=""><!--- Requires the Add/Edit finding notification to responsible person and closure email notification option to remain checked. --->
<cfparam name="ToList" default="">
<cfparam name="CCList" default="">
<cfparam name="CLosureVerificationlabelForEmails" default="verification">
<cfparam name="CAPA_SpecialRight" default="Closure Verifier">
<cfparam name="closureVerificationLabel" default="closure verification">
<cfparam name="RiskCategory_EditClosureDueDate" default="false">
<cfparam name="ClosureDueDateLabel" default="Closure Due Date">
<cfparam name="ltv_verifier" default="">
<cfparam name="form.replication" default="false">
<CFPARAM NAME="Form.AuditName" DEFAULT="">
<cfparam name="Form.coResponPerson" default="">
<cfparam name="coResponPersonValidation" default="12/14/2017">
<CFSET Lookup = CreateObject("component", "#Library.CFC.DotPath#.lookup.audit").init(odbc)>

<CFSET Translator = CreateObject("component", "#Library.CFC.DotPath#.TranslationPage").init(GroupPage="audaction.cfm")>

<cfset variables.extension = new "#request.library.cfc.dotpath#.extensions.extensions"().extInit(odbc,"ATS")>
<cfset variables.embedCustomTagger = go.getSetupVar(var="customTaggerEmbedAll", default=false)>

<!--- Get Auto Escalation variables --->
<cfset variables.autoEscalation_enable = api.go.getSetupVar(var="autoEscalation_enable",default=false)/>
<cfset variables.AutoEscalation_EscalationDays = go.getSetupVar(var="AutoEscalation_EscalationDays",default=5)/>

<!--- AJAX action for updating the closure verifier person from the status page --->
<CFIF FORM.Action IS "UpdateVerifyBy">
	<CFPARAM NAME="FORM.ID" DEFAULT="">
	<CFPARAM NAME="FORM.SiteID" DEFAULT="">
	<CFPARAM NAME="FORM.VerifyBy" DEFAULT="">

	<CFSET audit = createObject("component","#Request.Library.CFC.DotPath#.apps.ats.audit").init(ODBC="#ODBC#", AuditHome="#AuditHome#")>
	<CFSET qAudit = audit.getAudit(tblAuditID="#audit.getAuditID(SiteID="#FORM.SiteID#", ID="#FORM.ID#")#")>

	<CFIF Len(FORM.SiteID) EQ 0 OR NOT IsNumeric(FORM.SiteID) OR
			Len(FORM.ID) EQ 0 OR NOT IsNumeric(FORM.ID) OR
			Len(FORM.VerifyBy) EQ 0>
		<cfoutput>#Translator.Translate("There was a problem saving the new #ClosureVerificationByLabel#. Please reload this page and try again.")#</cfoutput><CFABORT>
	</CFIF>

	<CFIF !audit.isValidVerifier(verifier="#FORM.VerifyBy#", tblAuditID="#qAudit.tblAuditID#", CAPASpecialRightName="#CAPA_SpecialRight#", orgContactsOnly="#orgContactsOnly#")>
		<CFOUTPUT>#Translator.Translate("##USER## does not have the permissions required in order to be listed as the #ClosureVerificationByLabel#.", FORM.VerifyBy)#</CFOUTPUT><CFABORT>
	</CFIF>

	<CFPARAM NAME="BlockCloseVerify" DEFAULT="true">
	<CFIF BlockCloseVerify IS true AND FORM.VerifyBy IS NOT "" AND FORM.VerifyBy IS qAudit.ClosePerson>
		<CFOUTPUT>#Translator.Translate("The individual responsible for verification cannot be the same as the individual who closed the #LCase(REQUEST.ATSFindingName)#")#</CFOUTPUT><CFABORT>
	</CFIF>

	<CFSET audit.Update(VerifyBy="#FORM.VerifyBy#", TblAuditID="#qAudit.tblAuditID#", EmailRP="On") />

	<CFCONTENT RESET="true"><CFABORT>
</CFIF>

<!--- AJAX action for updating the responsible person from the status page --->
<CFIF FORM.Action IS "UpdateRespPerson">
    <CFPARAM NAME="FORM.ID" DEFAULT="">
    <CFPARAM NAME="FORM.SiteID" DEFAULT="">
    <CFPARAM NAME="FORM.ResponPerson" DEFAULT="">
 
    <CFSET audit = createObject("component","#Request.Library.CFC.DotPath#.apps.ats.audit").init(ODBC="#ODBC#", AuditHome="#AuditHome#")>
    <CFSET qAudit = audit.getAudit(tblAuditID="#audit.getAuditID(SiteID="#FORM.SiteID#", ID="#FORM.ID#")#")>
 
    <CFIF Len(FORM.SiteID) EQ 0 OR NOT IsNumeric(FORM.SiteID) OR
            Len(FORM.ID) EQ 0 OR NOT IsNumeric(FORM.ID) OR
            Len(FORM.ResponPerson) EQ 0>
        <cfoutput>#Translator.Translate("There was a problem saving the new Responsible Person. Please reload this page and try again.")#</cfoutput><CFABORT>
    </CFIF>
 
    <CFSET updateStatus = audit.Update(ResponPerson="#FORM.ResponPerson#", TblAuditID="#qAudit.tblAuditID#", EmailRP="On") />
    
    <CFIF updateStatus is false>
        <cfoutput>#Translator.Translate("There was a problem saving the new Responsible Person. Please reload this page and try again.")#</cfoutput><CFABORT>
    </CFIF>
 
    <CFCONTENT RESET="true"><CFABORT> 
</CFIF>

<CFIF FORM.ACtion IS "ExportCI">
	<CFSET audit = createObject("component","#Request.Library.CFC.DotPath#.apps.ats.audit").init(ODBC="#ODBC#", AuditHome="#AuditHome#")>
	<CFSET qAudit = audit.getAudit(tblAuditID="#audit.getAuditID(SiteID="#FORM.SiteID#", ID="#FORM.ID#")#")>

	<CFIF ListFindNoCase(CI_LOCK_ATS_FindingType, qAudit.FindingType) NEQ 0>
		<CFPARAM NAME="CI_LOCK_ATS_RCAlist" DEFAULT="">
		<CFPARAM NAME="CI_LOCK_ATS_Source" DEFAULT="Audit Finding/ATS">

		<CFSET AnalysisType = "NCA">
		<CFSET PrioritizationLevel = "Medium">
		<CFIF ListFindNoCase(CI_LOCK_ATS_RCAlist, qAudit.FindingType) NEQ 0>
			<CFSET AnalysisType = "RCA">
			<CFSET PrioritizationLevel = "High">
		</CFIF>
		<CFPARAM NAME="iAccessID" DEFAULT="#request.user.accessid#">
		<CFPARAM NAME="sAccessName" DEFAULT="#request.user.AccessName#">
		<!--- if there is not a valid user session use the variable passed in --->
		<CFIF iAccessID IS "">
			<CFPARAM NAME="FORM.init" DEFAULT="">
			<CFIF FORM.init IS NOT "">
				<CFSET iAccessID = VAL(FORM.init)>
				<CFSET qGetContact = Lookup.getLtbContact(init)>
				<CFSET sAccessName = qGetContact.Contact_name>
			</CFIF>
		</CFIF>
		<CFIF iAccessID IS NOT "">
			<cfset deptid=""/>
			<cfif qAudit.coe is not "">
				<cfquery name="queryCOE" DATASOURCE="#ODBC#">
					SELECT deptid
					FROM ltbCOE WITH (NOLOCK)
					WHERE OrgName=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#qAudit.orgname#">
					AND Location=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#qAudit.Location#">
					AND COE=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#qAudit.COE#">
				</cfquery>
				<cfset deptid=queryCOE.deptid/>
			</cfif>
			<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#appintegration\apptoci_email.cfm"
				odbc="#odbc#"
				action="added"
				acessid="#iAccessID#"
				accessname="#sAccessName#"
				siteid="#qAudit.SiteID#"
				initiator="#iAccessID#" <!---Responsible Person--->
				issuetype="Internal" <!--- *internal if from ATS* --->
				sourceofnonconformance="#CI_LOCK_ATS_Source#"  <!--- where it originates from --->
				ncoccurancedate="#qAudit.AuditDate#"<!--- finding date --->
				analysisinitiateddate="#dateformat(now(),"mm/dd/yyyy")#" <!--- today --->
				correctiveaction="#qAudit.CorrectiveAction#" <!--- finding corrective action --->
				containmentaction="#qAudit.CorrectiveAction#" <!--- finding corrective action --->
<!--- 				problemsignificance="#Replace(qAudit.Description, "|", " ", "ALL")#"  finding description --->
                Ncdetails="#qAudit.Description#"
				analysisleader="#qAudit.ResponPerson#"  <!--- responsible person --->
				analysistype="#AnalysisType#"  <!--- Major or Minor = NCA or RCA --->
				prioritizationlevel="#PrioritizationLevel#" <!--- high or medium --->
				problemoccurancedate="#qAudit.AuditDate#"<!--- finding date --->
				analysischeckpoint="0"
				status="0"
				archive="0"
				sip="1"
				reftype="ATS"
				refid="#qAudit.ID#"
				deptid="#deptid#"
				subcoeid="#qAudit.SubCOEID#"
				/>

			<cfset getlink = createObject("component","#Request.Library.CFC.DotPath#.apps.ci.ci").init(ODBC="#ODBC#").linktoci(apphome="#CIHome#",reftype="ats",refid="#qAudit.ID#",siteid="#qAudit.SiteID#")/>
			<BR>
			<div class="alert alert-warning" role="alert">
				Continuous Improvement Plan #<CFOUTPUT>#getlink.LinkID#</CFOUTPUT> has been initiated and this <CFOUTPUT>#Lcase(REQUEST.ATSFindingName)#</CFOUTPUT> is now locked until the NCA or RCA is complete.
				<BR>Please use this <span class="clickable" ONCLICK="window.open('<CFOUTPUT>#getlink.link#</CFOUTPUT>')">link</span>&nbsp;to finalize creation of the NCA or RCA in the Continuous Improvement Tool.</DIV>
			</div>
		<CFELSE>
			<BR>
			<div class="alert alert-danger" role="alert">
				Unable to initiate Continuous Improvement Plan, please login and try again.
			</DIV>
		</CFIF>
	</CFIF>

	<CFABORT>
</CFIF>

<!--- as a quick shortcut to "fix" this page from being an action page that can also display data, we store the form in a temp table
after submitting the data and then bring it back so it can then act as a view page, thus keeping double submits from occuring.
And YES, this is not particularly efficient since it is not running parts of the page multiple times, but to get it done without
rewritting the entire page, it was the quickest method --->
    <CFIF isDefined("URL.ActionID")>
        <!--- try/catch to contain when the page is refreshed and the temp table has been dropped
        --->
        <CFTRY>
            <CFQUERY NAME="qGetData" DATASOURCE="#ODBC#">
                DELETE FROM ####ATS_REDIRECT_DATA WHERE UpdateDate <= DateAdd(d, -2, getDate());
 
                SELECT Data
                FROM ####ATS_REDIRECT_DATA
                WHERE ID =  <CF_QUERYPARAM VALUE="#URL.ActionID#" CFSQLTYPE="CF_SQL_VARCHAR">
            </CFQUERY>
 
            <CFSET Data = DeserializeJSON(qGetData.Data)>
            <CFLOOP COLLECTION="#Data.URL#" ITEM="item">
                <CFIF Trim(item) IS NOT "">
                    <CFSET URL[item] = Data.URL[item]>
                </CFIF>
            </CFLOOP>
            <CFLOOP COLLECTION="#Data.FORM#" ITEM="item">
                <CFIF Trim(item) IS NOT "">
                    <CFSET FORM[item] = Data.FORM[item]>
                </CFIF>
            </CFLOOP>
 
            <CFCATCH>
			<cfif structKeyExists(url, "wtf")>
				<cfoutput><textarea>#EncodeForHTML(qGetData.Data)#</textarea></cfoutput>
				<cfdump var="#cfcatch#">
			  <cfabort>
			</cfif>
					
                <CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
                    You have accessed this page incorrectly. Please <A target="_top" HREF="audfinding.cfm" ONCLICK="history.go(-1); return false;">click here</A> to go back.
                </CFMODULE>
            </CFCATCH>
        </CFTRY>
    <CFELSE>
        
 
        <!--- save FORM and URL as they initially come in since things modify the scopes later on --->
        <CFSET REDIRECTDATA = {}>
        <CFSET REDIRECTDATA.URL = Duplicate(URL)>
        <CFSET REDIRECTDATA.FORM = Duplicate(FORM)>
        <!--- used for the code after the cflocation so it knows what had been going on --->
        <CFSET REDIRECTDATA.FORM.firstAction = REDIRECTDATA.FORM.Action>
        <CFSET REDIRECTDATA.FORM.Action = "View">
    </CFIF>


<cfif isdefined("url.frmsub")>
<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#formlostmsg.cfm" RequiredFields="PSTOKEN" SubmissionParam="frmSub" SearchForm NotifyAppLead="0">
</cfif>


<CFPARAM NAME="Request.User.AccessName" DEFAULT="">
<CFPARAM NAME="variables.AccessName" DEFAULT="#Request.User.AccessName#">
<cfif variables.AccessName EQ "" AND IsDefined("Form.ps_accessname")>
	<cfset AccessName = Form.ps_accessname>
	<CFSET Request.User.AccessName = AccessName>
</cfif>

<cfparam name="CopyAttachment" default="0">
<CFPARAM NAME="FORM.PDAMODE" DEFAULT="NO">
<CFPARAM NAME="FORM.OfflineMode" DEFAULT="NO">
<CFPARAM NAME="FORM.PCView" DEFAULT="NO">
<CFPARAM NAME="URL.PCView" DEFAULT="NO">
<CFPARAM NAME="Request.IsBlackBerry" DEFAULT="False">
<CFPARAM NAME="ExcludedATSAuditTypes" DEFAULT="">
<CFPARAM NAME="AuditGroupLimitType" DEFAULT="Framework Review">
<cfparam name="form.AppCaller" default="">
<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#PDA/PDACHECK.cfm">
<cfparam name="FORM.tblAuditID" default="">
<cfparam name="orglabel" default="Organization">
<cfparam name="sitelabel" default="Site">
<cfparam name="openerActionCallback" default=""> <!--- MB: 1/31/11: Opener function name to be called after action script executes --->

<CFPARAM NAME="FORM.OrigResponPerson" DEFAULT="">
<CFPARAM NAME="FORM.OrigCoResponPerson" DEFAULT="">
<CFPARAM NAME="Form.ID" DEFAULT="">
<CFPARAM NAME="Form.bAdded" DEFAULT="false">

<cfparam name="form.org" default="">
<cfparam name="form.location" default="">
<cfparam name="form.orgname" default="">
<cfparam name="form.siteID" default="">
<cfparam name="form.orgid" default="">

<cfparam name="form.firstAction" default="#form.Action#">
<cfif form.firstAction is "edit">
    <cfparam name="form.findingType"  default="">
    <cfset variables.scenario = "#form.firstAction#|#form.findingType#" />
<cfelse>
    <cfset variables.scenario = "add">
</cfif>
<cfset structAppend(variables, audit.getCustomSettings(variables.scenario),true)/>

<cfif isValid('integer',form.org)>
	<cfset org = form.org/>
</cfif>

<cfif StructKeyExists(url, "factoryID") and url.factoryid neq "">
	<cfset FactoryID = URL.factoryid>
<cfelse>
	<cfset FactoryID = "">
</cfif>


<!--- cfsavecontent Added so BlackBerry can submit back to form page and use a button, since it cannot use document.gobackform.submit() from a link -CJ 10/12/2005 --->
<cfsavecontent variable="thegobackform">
<cfoutput>
<cfset UseFormAction = "#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#audfinding.cfm?#CGI.QUERY_STRING#">
<span style="display:none;">
<form name="gobackform" method="post" action="#UseFormAction#">
<cfloop collection="#form#" item="data">
	<cftry>
	<cfset dataname = data>
	<cfset datavalue = Evaluate("form.#data#")>
		<input type="hidden" name="AutoFill_#EncodeForHTML(dataname)#" value="#EncodeForHTML(datavalue)#">
	<cfcatch>
	</cfcatch>
	</cftry>
</cfloop>
<cfif Request.IsBlackBerry EQ True>
<input type="submit" name="GoBack" value="Go Back">
</cfif>
<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#forminclude.cfm">
</form>
</span>
</cfoutput>
</cfsavecontent>

<cfif Request.IsBlackBerry NEQ True>
<cfoutput>
#thegobackform#
</cfoutput>
</cfif>

<cfparam name="PSValidDate" default="yes">
<CFPARAM NAME="Form.RefType" DEFAULT="">
<CFPARAM NAME="Form.RefID" DEFAULT="">
<CFPARAM NAME="Form.VerifyDate" DEFAULT="">


<!--- PowerSuite Sign-On --->
<cfset AppName = request.ATSAppName>
<cfparam name="SingleSignOn" default="">
<cfparam name="AccessLevel" default="">
<!--- <cfif SingleSignOn eq "Yes" AND FORM.OfflineMode IS "NO" And NOT IsDefined("FORM.SkipSSO")>
	<cfinclude template="#ContactsHome#Verification.cfm">
</cfif> --->
<!--- 07/14/03 --->

<CFPARAM NAME="org" DEFAULT = "">
<CFPARAM NAME="loc" DEFAULT = "">
<CFPARAM NAME="OName" DEFAULT = "">
<cfparam name="ClassificationTracking" default="">
<cfparam name="CostTracking" default="">
<CFPARAM Name="CitationLabel" Default="#REQUEST.ATSFindingName# Citation(s)">
<cfparam name="AuditorLabel" default="Auditor/Contact">
<cfparam name="coResponPersonLabel" default="#request.atsfindingname# Co-Owner"/>
<CFPARAM Name="Form.CloseComment" Default="">
<CFPARAM Name="Form.ClosePerson" Default="">
<CFPARAM Name="Form.ContactPhone" Default="">
<CFPARAM Name="Form.ContactPerson" Default="">
<CFPARAM Name="Form.NumItems" Default="1">
<CFPARAM Name="Form.Bldg" Default="">
<CFPARAM Name="Form.COE" Default="">
<CFPARAM Name="Form.FINDINGTYPE" Default="">
<CFPARAM Name="Form.Description" Default="">

<CFPARAM Name="Form.AuditType" Default="">
<CFPARAM Name="AuditActionTypeLabel" Default="Audit/Action Type">

<CFPARAM Name="Form.RepeatItem" Default="0">
<cfparam name="WorkstationLabel" default="Workstation"> <!--- Workstation label --->
<cfparam name="BuildingLabel" default="Building"> <!--- Building label --->
<cfparam name="ReferenceLabel" default="#REQUEST.ATSFindingName# Reference"> <!--- Finding Reference label --->
<CFPARAM NAME="CenterDeptLabel" DEFAULT="Center/Dept">
<CFPARAM NAME="Form.Category" DEFAULT="">
<CFPARAM NAME="FORM.subCat" DEFAULT="">
<CFPARAM NAME="Form.ClosureList" DEFAULT="">
<CFPARAM NAME="Form.Status" DEFAULT="">
<CFPARAM NAME="Form.AuditDate" DEFAULT="">
<CFPARAM NAME="Form.RiskCategory" DEFAULT="">
<CFPARAM NAME="Form.ExternalSubmit" DEFAULT="0">
<CFPARAM NAME="Effectivity2" DEFAULT="">
<CFPARAM NAME="RiskCategoryEnabled" DEFAULT="false">
<!--- 7/19/11 --->
<CFPARAM NAME="Form.Effectivity" DEFAULT=""> <!--- set default effectivity so ats upload template w/capa findings doesnt throw an error --->
<CFPARAM NAME="FiveWhyLabel" default="">

<!--- 9/14/15 GPS Variables --->
<CFPARAM NAME="variables.auditgpson" DEFAULT="0">
<CFPARAM NAME="GPSDisplay" DEFAULT="">
<CFPARAM NAME="GPSLabel" DEFAULT="GPS Coordinates">
<CFPARAM NAME="FORM.GPSlong" DEFAULT="">
<CFPARAM NAME="FORM.GPSlat" DEFAULT="">
<CFPARAM NAME="FORM.GPScomments" DEFAULT="">

<CFSET GPSLabel = Translator.Translate(GPSLabel)>

<!--- Force Email to be sent --->
<CFIF Form.ExternalSubmit EQ 1>
	<CFSET FORM.EmailRP = true>
</CFIF>


<CFIF ListFirst(Form.RiskCategory, "|") IS NOT "">
	<!--- risk category needs to be tied to a closure list item in order to override the value --->
	<cfquery name="ltbEnvQuery" datasource="#odbc#">
		select * from ltbEnvPriority with (nolock)
		where [EnvPriorityID] = <cf_queryparam cfsqltype="cf_sql_bigint" value="#val(listfirst(form.riskcategory, "|"))#">
	</cfquery>

	<cfif ltbEnvQuery.recordCount gt 0 and len(trim(ltbEnvQuery.ClosureID))>
		<!--- override the closurelist just to be sure --->
		<CFQUERY NAME="qGetClosureList" DATASOURCE="#ODBC#">
			SELECT Closure
			FROM ltbClosure WITH (NOLOCK)
			WHERE ClosureSortNo = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_BIGINT" VALUE="#VAL(ListFirst(Form.RiskCategory, "|"))#">
		</CFQUERY>
		<CFIF qGetClosureList.Closure IS NOT "">
			<CFSET FORM.ClosureList = qGetClosureList.Closure>
		</CFIF>
	</cfif>
</CFIF>
<cfif RiskCategoryEnabled>
	<CFSET Form.RiskCategory = ListLast(Form.RiskCategory, "|")>
<cfelse>
	<CFSET Form.RiskCategory = ListRest(Form.RiskCategory, "|")>
</cfif>
<CFIF ListLen(Form.COE, "|") EQ 2 AND NOT isDefined("FORM.SubCOEID")>
	<CFSET Form.SubCOEID = ListLast(Form.COE, "|")>
	<CFIF isNumeric(Form.SubCOEID)>
		<CFSET Form.COE = ListDeleteAt(Form.COE, 2, "|")>
	<CFELSE>
		<CFSET Form.SubCOEID = "">
	</CFIF>
</CFIF>

<CFSET ToList = "">
<CFSET CCList = "">

<cfif structKeyExists(FORM, "CC_Email")>
	<CFSET CCList = FORM.CC_Email>
</cfif>

<!--- correction and containment owners added to CC --->
<cfif structKeyExists(FORM, "CorrectionActionOwner") and FORM.CorrectionActionOwner is not "">
    <CFQUERY NAME="qGetCorrectionOwnerEmailCC" DATASOURCE="#ODBC#">
        SELECT CONTACT_EMAIL
        FROM ltbContact WITH (NOLOCK)
        WHERE Contact_Name = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.CorrectionActionOwner#">
    </CFQUERY>                
    <CFSET CCList = ListAppend(CCList, qGetCorrectionOwnerEmailCC.CONTACT_EMAIL)>
</cfif>

<cfif structKeyExists(FORM, "ContainmentActionOwner") and FORM.ContainmentActionOwner is not "">
    <CFQUERY NAME="qGetContainmentOwnerEmailCC" DATASOURCE="#ODBC#">
        SELECT CONTACT_EMAIL
        FROM ltbContact WITH (NOLOCK)
        WHERE Contact_Name = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.ContainmentActionOwner#">
    </CFQUERY>                
    <CFSET CCList = ListAppend(CCList, qGetContainmentOwnerEmailCC.CONTACT_EMAIL)>
</cfif>
<!--- end of addition --->

<cfparam name="SendEmailWhenFindingClosureDateModified" default="yes">

<!--- 08/16/02 --->
<CFPARAM NAME="VerifyBy" DEFAULT="">
<CFPARAM NAME="VerifyBy_Orig" DEFAULT="">

<CFPARAM NAME="VerifyComment" DEFAULT="">
<CFPARAM NAME="VerifyDate" DEFAULT="">

<cfparam name="RequireCorrectiveAction" default="Yes">

<cfquery name="qGetPriorityClass" datasource="#ODBC#">
	SELECT PriorityClass
	FROM ltbFindingType WITH (NOLOCK)
	WHERE FindingName = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.FindingType#">
</cfquery>

<cfif qGetPriorityClass.PriorityClass EQ "high">
	<cfset HighPriority = 1>
<cfelse>
	<cfset HighPriority = 0>
</cfif>

<CFPARAM NAME="RootCause1Column" DEFAULT="Optional_6">
<CFPARAM NAME="RootCause2Column" DEFAULT="Optional_7">
<CFPARAM NAME="FORM.RootCause1" DEFAULT="">
<CFPARAM NAME="FORM.RootCause2" DEFAULT="">
<CFPARAM NAME="SubCategoryColumn" default="Optional_8"><!--- This column is being used for the sub-category --->

<CFPARAM NAME="VerifyByDateMinimumDays" DEFAULT="0">

<CFPARAM NAME="FORM.firstAction" DEFAULT="#FORM.Action#">

<cfif form.ACTION EQ "Add">
	<!-- ACTION ADD -->
	<cfsavecontent variable="BACKLINK">
		<cfoutput>
			<A target="_top" HREF="audfinding.cfm?Org=#EncodeForURL(Org)#&Loc=#EncodeForURL(Loc)#&ID=#EncodeForURL(FORM.ID)#&FactoryID=#EncodeForURL(factoryid)#" ONCLICK="document.gobackform.submit(); return false;">click here</A>
		</cfoutput>
	</cfsavecontent>
<cfelse>
	<!-- NOT ACTION ADD -->
	<cfsavecontent variable="BACKLINK">
		<cfoutput>
		<A target="_top" HREF="audfinding.cfm?Org=#EncodeForURL(Org)#&Loc=#EncodeForURL(Loc)#&ID=#EncodeForURL(FORM.ID)#&FactoryID=#EncodeForURL(factoryid)#" ONCLICK="history.go(-1); return false;">click here</A>
		</cfoutput>
	</cfsavecontent>
</cfif>

<CFSET sAdditionalMessage = "">

<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#datechk.cfm" DateValue="#VerifyDate#" ReturnBoolean="YES">
<cfif PSValidDate EQ "NO" AND VerifyDate NEQ "">

<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
	<cfoutput><div id="error">The Verify Date, #VerifyDate#, is Invalid.</div> Please
	<cfif Request.IsBlackBerry EQ True>
		<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
		use the Go Back button to return to your input form and re-submit.
		#thegobackform#
	<!--- <cfelse>
	#BACKLINK#
	 to go back to the form and submit again after entering a correct date.--->
	</cfif>
	</cfoutput>
</CFMODULE>


</cfif>

<!--- 08/29/02 --->
<CFPARAM NAME="DaysBeforeReminder" DEFAULT="">

<!--- 10/20/03 --->
<CFPARAM NAME="Form.ResponPerson" DEFAULT="">
<!--- 10/27/03 --->
<CFPARAM NAME="bVerifyByDisabled" DEFAULT="false">

<!--- 11/06/02 --->
<CFSET ReOpenEmail = "no">
<!--- 04/20/04 --->
<CFPARAM NAME="FORM.SubCOEID" DEFAULT="">
<CFPARAM NAME="FORM.MULT_CC" DEFAULT="">
<CFPARAM NAME="Form.AuditType" DEFAULT="">

<CFPARAM NAME="RequireClosureVerificationOnAdd" DEFAULT="false">

<CFSET qAuditType = Lookup.getLtbAuditType(AT="#Form.AuditType#")>
<CFIF ListFindNoCase(ValueList(qAuditType.AuditName, chr(8)), FORM.AuditType, chr(8)) EQ 0>
	<CFSET Form.AuditType = "">
</CFIF>

<!--- 11/30/05 --->
<cfif Form.AuditType EQ "" AND CGI.REQUEST_METHOD EQ "POST">
	<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
		<cfoutput>
			<div id="error">#AuditTypeLabel# invalid.</div> Please
			<cfif Request.IsBlackBerry EQ True>
				<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
				use the Go Back button to return to your input form and re-submit.
				#thegobackform#
			<!--- <cfelse>
			#BACKLINK#
			 to go back to the form and submit again with a valid AuditType.--->
			</cfif>
		</cfoutput>
	</CFMODULE>
</cfif>

<!--- 04/03/03 --->
<CFPARAM NAME="VerifyByDate" DEFAULT="">
<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#datechk.cfm" DateValue="#VerifyByDate#" ReturnBoolean="YES">
<cfif PSValidDate EQ "NO" AND VerifyByDate NEQ "">
	<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
		<cfoutput><div id="error">Verify By Date "#VerifyByDate#" is Invalid.</div> Please
		<cfif Request.IsBlackBerry EQ True>
			<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
			use the Go Back button to return to your input form and re-submit.
			#thegobackform#
		<!--- <cfelse>
		#BACKLINK#
		 to go back to the form and submit again after correcting the date.--->
		</cfif>
		</cfoutput>
	</CFMODULE>
</cfif>
<CFPARAM NAME="VerifyByDatePreposition" DEFAULT="by">
<CFPARAM NAME="FORM.VerifyByDate" DEFAULT="">
<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#datechk.cfm" DateValue="#FORM.VerifyByDate#" ReturnBoolean="YES">
<cfif PSValidDate EQ "NO" AND FORM.VerifyByDate NEQ "">
	<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
		<cfoutput><div id="error">Verify By Date "#FORM.VerifyByDate#" is Invalid.</div> Please
		<cfif Request.IsBlackBerry EQ True>
			<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
			use the Go Back button to return to your input form and re-submit.
			#thegobackform#
		<!--- <cfelse>
		#BACKLINK#
		 to go back to the form and submit again after correcting the date.--->
		</cfif>
		</cfoutput>
	</CFMODULE>
</cfif>

<CFIF Org IS "">
	<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
	<div id="error">The Org is Invalid.</div>
	</CFMODULE>
</CFIF>

<!--- James Hunt: Dec 11, 2006  - Added to tollgate incorrect refid for calendar reftypes --->
<cfif reftype EQ "Calendar" AND Form.Action Is "Add">
	<CFSET refid = Int(VAL(refid))>
	<cfquery name="qGetValidRefId" datasource="#ODBC#">
		SELECT qryOrgSite.org AS orgname
		FROM TASK_REMINDER WITH (NOLOCK) INNER JOIN qryOrgSite WITH (NOLOCK)
			ON TASK_REMINDER.ORGNAME = qryOrgSite.ORGNAME
			AND TASK_REMINDER.LOCATION = qryOrgSite.Location
		WHERE  qryOrgSite.org = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_BIGINT" VALUE="#org#">
			AND TASK_REMINDER.LOCATION = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#loc#">
			AND dbo.TASK_REMINDER.TaskReminderID = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_BIGINT" VALUE="#VAL(refid)#">
	</cfquery>

	<cfif qGetValidRefId.RecordCount EQ 0>
		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<cfoutput><div id="error">The Reference ID "#RefID#" is Invalid.</div> Please
			<cfif Request.IsBlackBerry EQ True>
				<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
				use the Go Back button to return to your input form and re-submit.
				#thegobackform#
			<!--- <cfelse>
			#BACKLINK#
			 to go back to the form and submit again after correcting the task reminder date.--->
			</cfif>
			</cfoutput>
		</CFMODULE>
	</cfif>
</cfif>


<!--- 03/14/05 (moved 07/11/05) --->
<CFPARAM NAME="CAPALabel" DEFAULT="CAPA">
<CFPARAM NAME="ClosureCommentLabel" DEFAULT="Closure Comment">
<CFPARAM NAME="RiskCategoryLabel" DEFAULT="Risk Category">

<!--- 07/14/03 --->
<CFPARAM NAME="EncryptKey" DEFAULT="PowerSuite">

<!--- 06/17/04 --->
<CFPARAM NAME="FORM.InvestigationDetails" DEFAULT="">
<CFPARAM NAME="FORM.RootCause" DEFAULT="">
<CFPARAM NAME="FORM.EffectiveDate" DEFAULT="">
<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#DATECHK.cfm" DATEVALUE="#Form.EffectiveDate#" ReturnBoolean="YES">
<cfif PSValidDate EQ "NO" AND FORM.EffectiveDate NEQ "">
	<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
		<cfoutput><div id="error">Effective Date "#FORM.EffectiveDate#" is Invalid.</div> Please
		<cfif Request.IsBlackBerry EQ True>
			<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
			use the Go Back button to return to your input form and re-submit.
			#thegobackform#
		<!--- <cfelse>
		#BACKLINK#
		 to go back to the form and submit again after correcting the date.--->
		</cfif>
		</cfoutput>
	</CFMODULE>
</cfif>

<CFSET sCurrentURL = REQUEST.DOMAINPROTOCOL & Request.DomainURL & AuditHome>

<CFIF FORM.PDAMODE IS "yes">
	<CFPARAM NAME="Form.WorkStation" DEFAULT="">
	<CFPARAM NAME="Form.AuditName" DEFAULT="">
	<CFPARAM NAME="Form.REFTYPE" DEFAULT="">
	<CFPARAM NAME="Form.REFID" DEFAULT="">
</CFIF>
<CFPARAM NAME="Form.CITATION" DEFAULT="">

<!--- 05/03/04 --->
<CFPARAM NAME="AuditTypeFutureDayList" DEFAULT="">

<!--- 06/17/04 --->
<CFPARAM NAME="ATS_CAPA_ENABLED" DEFAULT="false">
<CFPARAM NAME="FORM.CAPARequired" DEFAULT="0">
<CFIF FORM.CAPARequired EQ 0>
	<CFSET ATS_CAPA_ENABLED = false>
</CFIF>
<CFPARAM NAME="ATS_CAPA_EFFECTIVITY_REQUIRED" DEFAULT="true">

<CFPARAM NAME="RootCauseColumn" DEFAULT="optional_3">
<CFPARAM NAME="EffectiveDateColumn" DEFAULT="optional_4">
<CFPARAM NAME="CAPARequiredColumn" DEFAULT="CAPARequired">
<CFPARAM NAME="EffectivityColumn1" DEFAULT="optional_1">
<CFPARAM NAME="EffectivityColumn2" DEFAULT="optional_5">
<CFPARAM NAME="FORM.optional_3" DEFAULT="">
<CFPARAM NAME="ATS_AUDIT_TRAIL" DEFAULT="false">



<CFIF ATS_CAPA_ENABLED IS true>
	<CFSET ATS_AUDIT_TRAIL = true>
</CFIF>

<cfset NewID = "">
<!--- Trying this so when this page is in the loop created by Sync Client,
the variable will be sure to change value.  -CJ 05/09/2005 --->

<CFIF ATS_CAPA_ENABLED IS true>
	<CFSET RequireClosureVerification = true>
</CFIF>

<CFIF FORM.OfflineMode IS "yes">
	<!--- verify that the AccessName is valid and get the access level --->
	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#OfflineMode/PDAMODE.cfm" ACTION="AccessNameVerify"
		ENCRYPTKEY="#EncryptKey#"
		ORGID="#ORG#"
		LOCATION="#Loc#"
		APPLICATION="#AppName#"
		ODBC="#ODBC#"></CFMODULE>

	<CFIF AccessLevel EQ 0 OR AccessLevel IS "">
		<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#PDA/PDASYNC.cfm" ACTION="Failure"></CFMODULE>
		<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#PDA/PDASync.cfm" ACTION="Show">
		<!--- deny ability to add Finding --->
		<CFSET DenyType = "Add #REQUEST.ATSFindingName#">
		<CFSET DenyLevel = 1>
		<CFSET DenyText = "To add a new #LCase(REQUEST.ATSFindingName)#, you must have <i>minimum</i> Level 1 permissions.">
		<CFSAVECONTENT VARIABLE="DenySignOn">
			<CFINCLUDE TEMPLATE="#ContactsHome#DenySignOn.cfm">
		</CFSAVECONTENT>
		<CFSET PATH = "/" & ListGetAt(CGI.SCRIPT_NAME, 1, "/") & "/" & ListGetAt(CGI.SCRIPT_NAME, 2, "/") & "/">
		<CFSET DenySignOn = ReplaceNoCase(DenySignOn, PATH, "#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##PATH#", "All")>

		<CFOUTPUT>
			#DenySignOn#
		</CFOUTPUT>
		</CFMODULE>
		<CFABORT>
	</CFIF>

</CFIF>

<!--- 03/25/04 --->
<CFPARAM NAME="FORM.CorrectiveAction" DEFAULT="">
<CFSET VerifyComment = Trim(VerifyComment)>
<CFSET FORM.Description = Trim(FORM.Description)>
<CFSET FORM.CorrectiveAction = Trim(FORM.CorrectiveAction)>
<CFSET FORM.CloseComment = Trim(FORM.CloseComment)>
<cftry>
	<cfif CGI.Server_Name CONTAINS "cincep09corpge">
		<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#MaliciousCodeStripper.cfm" input="#Form#" output="Form">
	</cfif>
	<cfcatch></cfcatch>
</cftry>


<cfif #ClassificationTracking# neq "">
	<cfset OrgPosition = #ListFind(ClassificationTracking,Org)#>
	<cfset CostTracking = #ListGetAt(ClassificationTracking,(OrgPosition+1))#>
</cfif>



<CFQUERY NAME="OrgQ" DATASOURCE="#ODBC#">
    SELECT OrgName, OrgPassword
	FROM Org  WITH (NOLOCK)
	WHERE ([TABLE]=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_BIGINT" VALUE="#org#">);
</CFQUERY>

<CFOUTPUT QUERY="OrgQ">
    <CFSET #OName# = #OrgName#>
</CFOUTPUT>

<!--- 04/20/04 --->
<CFSET SubCOEName = "">
<CFIF FORM.SubCOEID IS NOT "">
	<CFQUERY NAME="qGetSubCOE" DATASOURCE="#ODBC#">
		SELECT SubCOE, SubCOEID
		FROM ltbCOE_Sub WITH (NOLOCK)
			inner join ltbCOE (NOLOCK)
				on ltbCOE.deptid = ltbCOE_sub.deptid
				and ltbCOE.orgname = N'#oname#'
				and ltbCOE.location = N'#Loc#'
		WHERE SubCOEID = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_BIGINT" VALUE="#FORM.SubCOEID#">
	</CFQUERY>
	<CFSET SubCOEName = qGetSubCOE.SubCOE>
	<cfset FORM.SubCOEID = qGetSubCOE.SubCOEID>
</CFIF>


<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#SiteNotActive.cfm" AppName="#AppName#" location="#loc#" orgname="#oname#"
	icon="#Library.Images.URL#AppIcons/30x30/ats.gif"
	topbar="#Library.Images.URL#AppIcons/topbars/ats.gif"
	textstyle="color:navy;"
	weblink="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##CalendarHome#org.cfm?org=#org#"
	resources = "#AppName# Online Help;#CalendarHome#index.htm"
	odbc="#ODBC#"
	fieldname="ATSCC_ON"
	>


<CFSET sAction = FORM.firstAction>
<CFIF sAction IS "Delete">
	<CFSET  sAction = sAction & "d">
<CFELSE>
	<CFSET  sAction = sAction & "ed">
</CFIF>

<cfif FORM.PDAMODE NEQ "Yes" OR FORM.OfflineMode IS "YES">
	<cfset variables.title = "#REQUEST.ATSFindingName# #sAction#"/>
	<cfset variables.title &= structKeyExists(form, "id")&&isValid("integer", form.id)?": ###form.id#":''/>
	<cfmodule template="topbar.cfm" />
<cfelse>
	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#PDA/PDAtopBar.cfm" image="ats.gif"
			alt="#request.ATSAppName#"
			AppName = "#AppName#"/>
</cfif>
<cfset request.ATSAuditName = ""/><!--- don't remove --->
<!--- Start Capturing Data for PDA --->
<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#PDA/PDASync.cfm" ACTION="Show">

<!--- 09/19/03 --->
<CFIF (#ParameterExists(Form.Action)# is NOT "Yes" OR FORM.Action IS "") and (NOT isdefined("addsimilar"))>
	<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
	<cfoutput>
		<div id="error">
			Note: This form should be accessed to Add, Change or Delete an #REQUEST.ATSAuditName# #REQUEST.ATSFindingName#.<BR><BR>
			<A target="_top" HREF="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##ContactsHome#email.cfm">Contact</A> the business administrator for assistance.
		</div>
	</cfoutput>
	</CFMODULE>
</CFIF>

<CFIF Form.Action NEQ "Delete" and not IsDefined("AddSimilar") AND FORM.Action IS NOT "View">
<!--- 03/17/04 --->
	
	<CFIF ParameterExists(Form.Location) EQ "No"
		OR (Form.FINDINGTYPE EQ "" AND FindingTypeRequired)
		OR ParameterExists(Form.WorkStation) EQ "No"
		OR ParameterExists(Form.AuditName) NEQ "Yes"
		OR NOT isDefined("FORM.STATUS")>

	
<!--- 
		<cfdump var="#ParameterExists(Form.Location) EQ "No"
		OR (Form.FINDINGTYPE EQ "" AND FindingTypeRequired)
		OR ParameterExists(Form.WorkStation) EQ "No"
		OR ParameterExists(Form.AuditName) NEQ "Yes"
		OR NOT isDefined("FORM.STATUS")#"> --->


		<!--- 
		<cfdump var="#ParameterExists(Form.Location) EQ "No"#">
				<cfdump var="#Form.FINDINGTYPE EQ "" AND FindingTypeRequired#">
				<cfdump var="#ParameterExists(Form.WorkStation)#"> --->
		
		
		<!--- <cfdump var="#NOT isDefined("FORM.STATUS")#"> --->

		
		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<Font size=+1><div id="error">Sorry!<cfoutput>#Form.Category#</cfoutput> Network and/or browser difficulties are preventing your inputs from being received completely for processing!</div><p>
			<cfif Request.IsBlackBerry EQ True>
			<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
			<cfoutput>
			Please use the Go Back button to return to your input form and re-submit.
			#thegobackform#
			</cfoutput>
			<!--- <cfelse>
			<cfoutput>#BACKLINK# </cfoutput>
			 to go back and re-submit your input form.  Please note that you may need to reload your form page in order to submit your inputs successfully. ---></font><BR><BR>
			</cfif>
			<cfif Request.IsBlackBerry NEQ True>
			<cfoutput>
			<A target="_top" HREF="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##ContactsHome#email.cfm">Contact</A> the business administrator if this problem persists. Please be sure to use Microsoft Internet Explorer 5.0 or higher for best results.
			</cfoutput>
			</cfif>
		</CFMODULE>
	</CFIF>

	<CFIF Form.Category eq "">
		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<div id="error">A category is required.</div><p>
			<cfif Request.IsBlackBerry EQ True>
			<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
			<cfoutput>
			Please use the Go Back button to return to your input form and re-submit.
			#thegobackform#
			</cfoutput>
			<!--- <cfelse>
			<cfoutput>#BACKLINK# </cfoutput>
			 to go back and re-submit your input form.  Please note that you may need to reload your form page in order to submit your inputs successfully. ---></font><BR><BR>
			</cfif>
			<cfif Request.IsBlackBerry NEQ True>
			<cfoutput>
			<A target="_top" HREF="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##ContactsHome#email.cfm">Contact</A> the business administrator if this problem persists.  Please be sure to use Microsoft Internet Explorer 5.0 or higher for best results.
			</cfoutput>
			</cfif>
		</CFMODULE>
	</CFIF>

</CFIF>


<cfif URL.PCView IS "Yes" OR PDAMode NEQ True>
	<cfset UseFontTagStyle = "class='locHeader'">
	<cfset UseFontTagStyle2 = "class='orgHeader'">
	<cfset UseFontTagStyle3 = "class='largetext'">
<cfelse>
	<cfset UseFontTagStyle = "class='locHeader' SIZE=+3">
	<cfset UseFontTagStyle2 = "class='orgHeader' SIZE=+2">
	<cfset UseFontTagStyle3 = "class='largetext' SIZE=+3">
</cfif>

<CFOUTPUT>
<FONT #UseFontTagStyle#>
</CFOUTPUT>

<CFQUERY NAME="qGetSiteID" DATASOURCE="#ODBC#">
	SELECT SiteID
	FROM Site WITH (NOLOCK)
	WHERE OrgName=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
	AND Location=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Loc#">
</CFQUERY>
<CFSET SiteID = qGetSiteID.SiteID>

<cfif FactoryID eq "" and len(trim(SiteID))>
	<CFQUERY NAME="qGetFactoryID" datasource="#odbc#">
		SELECT isnull(textanswer,'') + ' - ' as FactoryID 
		FROM PROFILE_DATA WITH (NOLOCK) WHERE OPTIONID = #SiteID#
		and ITEMCODE =N'supplier_factoryid' 
	</cfquery>
	<cfset factoryid = qGetFactoryID.FactoryID>
</cfif>

<CFSET bHasSubCat = Lookup.getLtbCategorySub(siteid="#siteid#",archive=0,archiveCat=0).RecordCount NEQ 0>

<CFPARAM NAME="CI_actions_comp" DEFAULT="">
<CFIF CI_LOCK_ATS_FindingType IS NOT "" AND ListFindNoCase(CI_LOCK_ATS_FindingType, Form.FindingType) NEQ 0>
	<CFPARAM NAME="CIHome" DEFAULT="">

	<cfset CILink = createObject("component","#Request.Library.CFC.DotPath#.apps.ci.ci").init(ODBC="#ODBC#").linktoci(apphome="#CIHome#",reftype="ats",refid="#id#",siteid="#SiteID#")/>
	<cfset CI_actions_comp = CILINK.ACTIONSCOMPLETE>
</CFIF>


<!--- param all the fields --->
<CFMODULE 
    TEMPLATE = "#Request.Library.CustomTags.VirtualPath#additionalFields.cfm"
	FIELDS   = "additionalATSFields"
	orgid    = "#form.orgid#"
	odbc     = "#odbc#" 
	TYPE     = "DEFAULT" 
	SCOPE    = "#form#"/>

<cfset additionalSharedRows = ""><!--- ge aviation qlty added 2 fields that need to be adjacent to one other for readability as users can also add more of the same fields. ALL FIELD NAMES SHOULD END IN 3 NUMERIC DIGITS (typically padded zeros created with custom.js) Logic may need improvements --->
<CFIF isDefined("additionalATSFields") AND isStruct(additionalATSFields)>
	<cfloop list="#structKeyList(additionalATSFields)#" index="additionalGroup">
		<cfloop array="#additionalATSFields[additionalGroup]#" index="Field">
			<cfif structKeyExists(Field, "shareRowWith") && !ListFindNoCase(additionalSharedRows,Field.shareRowWith)>
				<cfset additionalSharedRows = ListAppend(additionalSharedRows,Field.shareRowWith)>
			</cfif>
		</cfloop>
	</cfloop>
</CFIF>

<CFIF Form.Action Is "Add" OR Form.Action Is "Edit" OR Form.Action Is "View">
	<CFSET ValidNum = "Yes">
	<CFIF IsNumeric(Form.NumItems) Is "Yes">
	  <CFIF Form.NumItems LTE 0><CFSET ValidNum = "No"></CFIF>
	</CFIF>

	<CFIF ATS_CAPA_EFFECTIVITY_REQUIRED IS true AND (FORM.CAPARequired NEQ 0) AND (FORM.Status IS "Closed") AND (FORM.EffectiveDate EQ "")>
		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<CFOUTPUT>
			<div id="error">The #EffectivityDateLabel# <b>can not be blank</b> if #REQUEST.ATSFindingName# Status is <i>closed</i> and #LCase(REQUEST.ATSAuditName)# is subject to #CAPALabel# requirements.</div>
			<cfif Request.IsBlackBerry EQ True>
			<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
			Please use the Go Back button to return to your input form and re-submit.
			#thegobackform#
			<!--- <cfelse>
			#BACKLINK#
			 to go back and correct the input. --->
			</cfif>
			</CFOUTPUT>
		</CFMODULE>
	</CFIF>
	<CFPARAM NAME="FORM.ignoreCILock" DEFAULT="false">
	<CFIF CI_LOCK_ATS_FindingType IS NOT "" AND ListFindNoCase(CI_LOCK_ATS_FindingType, Form.FindingType) NEQ 0 AND FORM.Status IS "Closed" AND NOT isDefined("FORM.CFCClose") AND CILink.actionsComplete IS "" AND FORM.ignoreCILock IS false AND (ByPassExport IS FALSE OR FORM.EXPCIMandatory contains "CIExport")>
		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<CFOUTPUT>
			<div id="error">You cannot close the #Lcase(REQUEST.ATSFindingName)# because the #REQUEST.ATSFindingName# Type requires the #Lcase(REQUEST.ATSFindingName)# be exported to Continuous Improvement.</div>
			<cfif Request.IsBlackBerry EQ True>
			<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
			Please use the Go Back button to return to your input form and re-submit.
			#thegobackform#
			<!--- <cfelse>
			#BACKLINK#
			 to go back and correct the input. --->
			</cfif>
			</CFOUTPUT>
		</CFMODULE>
	</CFIF>

	<CFIF ATS_CAPA_EFFECTIVITY_REQUIRED IS true AND (FORM.CAPARequired NEQ 0) AND (FORM.Status IS "Closed") AND (FORM.EffectiveDate EQ "")>
		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<CFOUTPUT>
			<div id="error">The #EffectivityDateLabel# <b>can not be blank</b> if #REQUEST.ATSFindingName# Status is <i>closed</i> and #LCase(REQUEST.ATSAuditName)# is subject to #CAPALabel# requirements.</div>
			<cfif Request.IsBlackBerry EQ True>
			<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
			Please use the Go Back button to return to your input form and re-submit.
			#thegobackform#
			<!--- <cfelse>
			#BACKLINK#
			 to go back and correct the input. --->
			</cfif>
			</CFOUTPUT>
		</CFMODULE>
	</CFIF>

	<CFIF ((IsNumeric(Form.NumItems) Is Not "Yes") OR (ValidNum Is Not "Yes")) and (numberOfItemsLabel is not "")>
		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<CFOUTPUT>
			<div id="error">The Number of Individual Items associated with this #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# must be a whole number > 0.</div>
			<cfif Request.IsBlackBerry EQ True>
			<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
			Please use the Go Back button to return to your input form and re-submit.
			#thegobackform#
			<!--- <cfelse>
			#BACKLINK#
			 to go back and correct the input. --->
			</cfif>
			</CFOUTPUT>
		</CFMODULE>
	</CFIF>

	<CFIF (IsDate(Form.AuditDate) Is NOT "Yes")>
		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<CFOUTPUT>
			<div id="error">#REQUEST.ATSFindingName# Date #Form.AuditDate# is Invalid.</div>
			<cfif Request.IsBlackBerry EQ True>
			<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
			use the Go Back button to return to your input form and re-submit.
			#thegobackform#
			<!--- <cfelse>
			#BACKLINK#
			 to go back and correct the input. --->
			</cfif>
			</CFOUTPUT>
		</CFMODULE>
	</CFIF>
	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#datechk.cfm" DateValue="#Form.AuditDate#" ReturnBoolean="YES">
	<cfif PSValidDate EQ "NO" AND FORM.AuditDate NEQ "">
		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<cfoutput><div id="error">#REQUEST.ATSFindingName# Date #FORM.AuditDate# is Invalid.</div> Please
			<cfif Request.IsBlackBerry EQ True>
			<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
			use the Go Back button to return to your input form and re-submit.
			#thegobackform#
			<!--- <cfelse>
			#BACKLINK#
			 to go back to the form and submit again after entering a correct the date. --->
			</cfif>
			</cfoutput>
		</CFMODULE>
	</cfif>

	<CFIF Form.ClosureList EQ "">
		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<CFOUTPUT>
			<div id="error">Closure Category must be selected.</div>
			<cfif Request.IsBlackBerry EQ True>
			<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
			Please use the Go Back button to return to your input form and re-submit.
			#thegobackform#
			<!--- <cfelse>
			#BACKLINK#
			 to go back and correct the input. --->
			</cfif>
			</CFOUTPUT>
		</CFMODULE>
	</CFIF>

	<!--- 05/03/04 --->
	<CFSET iFutureDays = 1>
	<CFTRY>
		<CFLOOP INDEX="AuditTypeItem" LIST="#AuditTypeFutureDayList#" DELIMITERS="|">
			<CFSET sAuditType = ListGetAt(AuditTypeItem, 1)>
			<CFIF isDefined("FORM.AuditType") AND sAuditType IS FORM.AuditType>
				<CFSET iFutureDays = ListGetAt(AuditTypeItem, 2)>
			</CFIF>
		</CFLOOP>
		<CFCATCH>

			<CFMAIL TO="ActionTrackingSystem.PM@gensuitellc.com" FROM="ActionTrackingSystem.PM@gensuitellc.com"
				 SUBJECT="audaction.cfm - Failure" TYPE="HTML">
			 	 <CFDUMP VAR="#CGI#">
				 <CFDUMP VAR="#CFCATCH#">
			</CFMAIL>
		</CFCATCH>
	</CFTRY>

	<CFIF Form.AuditDate GT DateAdd('d', iFutureDays, Now())>
		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<CFOUTPUT>
			<div id="error">#REQUEST.ATSFindingName# Date, <i>#Form.AuditDate#</i>, cannot be later than <i>#DateFormat(DateAdd('d', iFutureDays, Now()),"dd-mmm-yyyy")#</i></div>
			<cfif Request.IsBlackBerry EQ True>
			<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
			Please use the Go Back button to return to your input form and re-submit.
			#thegobackform#
			<!--- <cfelse>
			#BACKLINK#
			 to go back and correct the input. --->
			</cfif>
			</CFOUTPUT>
		</CFMODULE>
	</CFIF>

	<cfquery name="queryCOE" DATASOURCE="#ODBC#">
		SELECT COE,deptid
		FROM ltbCOE WITH (NOLOCK)
		WHERE OrgName=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
		 AND Location=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.Location#">
		 AND COE=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.COE#">
	</cfquery>
	<cfif queryCOE.RecordCount EQ 0>
		<cfset Form.COE = "">
	</cfif>

	<cfquery name="queryBldg" DATASOURCE="#ODBC#">
		SELECT Bldg
		FROM ltbBuilding WITH (NOLOCK)
		WHERE OrgName=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
			AND Location=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.Location#">
		 	AND Bldg=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.Bldg#">
	</cfquery>
	<cfif queryBldg.RecordCount EQ 0>
		<cfset Form.Bldg = "">
	</cfif>

	<CFQUERY NAME="qGetWorkstation" DATASOURCE="#ODBC#">
		SELECT WorkstationName
		FROM ltbWorkstation WITH (NOLOCK)
		WHERE SiteID = #VAL(SiteID)#
			AND WorkstationName = N'#FORM.Workstation#'
	</CFQUERY>
	<cfif qGetWorkstation.WorkstationName IS "">
		<cfset FORM.Workstation = "">
	</cfif>

	<cfquery name="queryRP" DATASOURCE="#ODBC#">
		SELECT Contact_Name
		FROM ltbContact WITH (NOLOCK)
		<cfif isValid("integer",form.ResponPerson)>
			WHERE Contactid=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_INTEGER" VALUE="#Form.ResponPerson#">
		<cfelse>
			WHERE Contact_Name=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.ResponPerson#">
		</cfif>
	</cfquery>
	<cfset Form.ResponPerson = queryRP.Contact_name/>
	<cfif queryRP.RecordCount EQ 0 AND (Form.ID IS "" OR (Form.ResponPerson IS NOT FORM.OrigResponPerson AND Form.ID IS NOT ""))>
		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<!--- 10/20/03 --->
			<CFIF Form.ResponPerson IS "">
				<div id="error">A Responsible Person must be selected!</div>
			<CFELSE>
				<div id="error">The Responsible Person no longer exists in the Contacts Database.</div>
			</CFIF>
			<cfif Request.IsBlackBerry EQ True>
			<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
			Please use the Go Back button to return to your input form and re-submit.
			<cfoutput>#thegobackform#</cfoutput>
			<!--- <cfelse>
			<cfoutput>#BACKLINK# </cfoutput>
			 to go back and correct the input(s). --->
			</cfif>
			</font>
		</CFMODULE>
	</cfif>
	<cfif variables.showCoResponPerson and structKeyExists(form, "AuditDate") and (dateCompare(form.AuditDate, coResponPersonValidation) GT 0)>
		<cfif Form.coResponPerson eq Form.ResponPerson>
			<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
				<div id="error">A <cfoutput>#coResponPersonLabel#</cfoutput> must be different than Responsible Person!</div>
				<cfif Request.IsBlackBerry EQ True>
				<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
					Please use the Go Back button to return to your input form and re-submit.
					<cfoutput>#thegobackform#</cfoutput>
				<!--- <cfelse>
					<cfoutput>#BACKLINK# </cfoutput>
				 	to go back and correct the input(s). --->
				</cfif>
				</font>
			</CFMODULE>
		</cfif>
		<cfquery name="queryCoRP" DATASOURCE="#ODBC#">
			SELECT Contact_Name
			FROM ltbContact WITH (NOLOCK)
			<cfif isValid("integer",form.coResponPerson)>
				WHERE Contactid=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_INTEGER" VALUE="#Form.coResponPerson#">
			<cfelse>
				WHERE Contact_Name=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.coResponPerson#">
			</cfif>
		</cfquery>
		<cfset Form.coResponPerson = queryCoRP.Contact_name/>
		<cfif variables.isCoResponPersonRequired eq true>
		<cfif querycoRP.RecordCount EQ 0 AND (Form.ID IS "" OR (Form.CoResponPerson IS NOT FORM.OrigCoResponPerson AND Form.ID IS NOT ""))>
			<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
				<!--- 10/20/03 --->
				<!--- Adding validation for Finding Co-Owner to verified Field is required after Dic-15 which is the relase date for this new feature --->
				<cfoutput>
				<CFIF Form.coResponPerson IS "" and (#dateCompare(form.AuditDate, '12/15/2017')# GT 0)>
					<div id="error">A #coResponPersonLabel# must be selected!</div>
				<CFELSE>
					<div id="error">The #coResponPersonLabel# no longer exists in the Contacts Database.</div>
				</CFIF>
				</cfoutput>
				<cfif Request.IsBlackBerry EQ True>
				<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
				Please use the Go Back button to return to your input form and re-submit.
				<cfoutput>#thegobackform#</cfoutput>
				<!--- <cfelse>
				<cfoutput>#BACKLINK# </cfoutput>
				 to go back and correct the input(s). --->
				</cfif>
				</font>
			</CFMODULE>
		</cfif>
	</cfif>
	</cfif>

	<!--- moved here 6/19/98 --->
    <CFIF ParameterExists(Form.Status) Is "Yes">
		<CFIF (Form.Status Is "Closed")>
			<CFIF (IsDate(Form.CloseDate) Is NOT "Yes")>
				<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
					<CFOUTPUT>
					<div id="error">A valid Closed Date must be provided for a Closed #REQUEST.ATSAuditName# #REQUEST.ATSFindingName#.</div>
					<cfif Request.IsBlackBerry EQ True>
					<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
					Please use the Go Back button to return to your input form and re-submit.
					#thegobackform#
					<!--- <cfelse>
					#BACKLINK#
					to go back and correct the input(s). --->
					</cfif>
					</CFOUTPUT>
				</CFMODULE>
			</CFIF>
			<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#datechk.cfm" DateValue="#Form.CloseDate#" ReturnBoolean="YES">
			<cfif PSValidDate EQ "NO" AND FORM.CloseDate NEQ "">
				<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
					<cfoutput>
					<div id="error">Close date "#Form.CloseDate#" is Invalid.</div> Please
					<cfif Request.IsBlackBerry EQ True>
					<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
					Please use the Go Back button to return to your input form and re-submit.
					#thegobackform#
					<!--- <cfelse>
					#BACKLINK#
					 to go back to the form and submit again after correcting the date. --->
					</cfif>
					</cfoutput>
				</CFMODULE>
			</cfif>

			<CFIF Form.CloseDate GT DateAdd('d', 1, Now())>
				<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
					<CFOUTPUT>
					<div id="error">Closed Date, <i>#DateFormat(Form.CloseDate, "dd-mmm-yyyy")#</i>, must be earlier than or equal to <i>"#DateFormat(DateAdd('d',1,Now()), "dd-mmm-yyyy")#"</i>!</div>
					<cfif Request.IsBlackBerry EQ True>
					<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
					Please use the Go Back button to return to your input form and re-submit.
					#thegobackform#
					<!--- <cfelse>
					#BACKLINK#
					 to go back and correct the input. --->
					</cfif>
					</CFOUTPUT>
				</CFMODULE>
			</CFIF>
			<!--- added 6/17/98 --->
			<CFIF #DateCompare(Form.AuditDate, Form.CloseDate)# GT 0>
				<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
					<CFOUTPUT>
					<div id="error">Closed Date for the #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# must be equal to or later than the #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# date!</div><BR>
					<cfif Request.IsBlackBerry EQ True>
					<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
					Please use the Go Back button to return to your input form and re-submit.
					#thegobackform#
					<!--- <cfelse>
					#BACKLINK#
					 to go back and correct the input. --->
					</cfif>
					</CFOUTPUT>
				</CFMODULE>
			</CFIF>
			<!--- Added 4/1/02 in case of JS bypass --->
			<CFIF Trim(Form.CloseComment) EQ "" or Form.ClosePerson EQ "">
				<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
					<CFOUTPUT>
					<div id="error">A Closure Comment must be provided and a Closed By person must be specified for a Closed #REQUEST.ATSAuditName# #REQUEST.ATSFindingName#.</div>
					<cfif Request.IsBlackBerry EQ True>
					<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
					Please use the Go Back button to return to your input form and re-submit.
					#thegobackform#
					<!--- <cfelse>
					#BACKLINK#
					 to go back and correct the input(s). --->
					</cfif>
					</CFOUTPUT>
				</CFMODULE>
			</CFIF>

			<!--- 11/04/02 --->
			<CFIF FORM.Status IS "Closed" AND FORM.VerifyDate IS NOT "" AND isDate(FORM.VerifyDate)
			 AND DateCompare(Form.CloseDate, FORM.VerifyDate) EQ 1>
			 	<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
					<CFOUTPUT>
					<div id="error">Closure Verification Date for the #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# must be equal to or later than the #DateClosedLabel#!</div>
					<cfif Request.IsBlackBerry EQ True>
					<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
					Please use the Go Back button to return to your input form and re-submit.
					#thegobackform#
					<!--- <cfelse>
					#BACKLINK#
					to go back and correct the input. --->
					</cfif>
					</CFOUTPUT>
				</CFMODULE>
			</CFIF>
			<!--- ensure the VerifyDate is the correct number of days in the future --->
			<CFIF VAL(VerifyByDateMinimumDays) GT 0 AND FORM.VerifyByDate IS NOT "" AND DateCompare(FORM.VerifyByDate, DateAdd("d", Form.CloseDate, VerifyByDateMinimumDays)) EQ -1>
				<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
					<CFOUTPUT>
					<div id="error">Scheduled Verification Date for the must be #VAL(VerifyByDateMinimumDays)# days after the  #DateClosedLabel#!</div>
					<cfif Request.IsBlackBerry EQ True>
					<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
					Please use the Go Back button to return to your input form and re-submit.
					#thegobackform#
					<!--- <cfelse>
					#BACKLINK#
					to go back and correct the input. --->
					</cfif>
					</CFOUTPUT>
				</CFMODULE>
			</CFIF>

		</CFIF>
    </CFIF>

	<CFSAVECONTENT VARIABLE="Message">
						<CFOUTPUT>
						<cfif Request.IsBlackBerry EQ True>
						<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
						Please use the Go Back button to return to your input form and re-submit.
						#thegobackform#
						<!--- <cfelse>
						#BACKLINK#
						to go back and correct the input. --->
						</cfif>
						</CFOUTPUT>
	</CFSAVECONTENT>

	<!--- verify that the required fields where submited --->
	<CFMODULE 
		TEMPLATE = "#Request.Library.CustomTags.VirtualPath#additionalFields.cfm"
		FIELDS   = "additionalATSFields" 
		TYPE     = "VALIDATE" 
		odbc     = "#odbc#" 
		orgid    = "#form.orgid#"
		VALIDATIONTYPE = "SCOPE" 
		SCOPE    = "#FORM#"
		MESSAGE  = "#MESSAGE#"/>

</CFIF>

<CFPARAM Name="OrgContactsOnly" Default="Yes">

<!--- 10/05/04 --->
<CFPARAM NAME="FORM.VerifyBy" DEFAULT="">
<CFIF FORM.Action IS NOT "Delete" AND FORM.Action IS NOT "View">

	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#GRANTPERMISSIONS.cfm"
		CONTACTNAME="#FORM.VerifyBy#"
		APPNAME="#request.ATSAppName#"
		LEVEL="1"
		ORGNAME="#OName#"
		LOCATION="#Loc#">

	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#GRANTPERMISSIONS.cfm"
		CONTACTNAME="#FORM.ResponPerson#"
		APPNAME="#request.ATSAppName#"
		LEVEL="1"
		ORGNAME="#OName#"
		LOCATION="#Loc#">
	<cfif variables.showCoResponPerson and form.coResponPerson neq "">
		<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#GRANTPERMISSIONS.cfm"
			CONTACTNAME="#FORM.coResponPerson#"
			APPNAME="#request.ATSAppName#"
			LEVEL="1"
			ORGNAME="#OName#"
			LOCATION="#Loc#">
	</cfif>


	<CFIF FORM.VerifyBy IS NOT "">
		<CFIF !audit.isValidVerifier(verifier="#FORM.VerifyBy#", siteID="#SiteID#", auditType="#FORM.auditType#", CAPASpecialRightName="#CAPA_SpecialRight#", orgContactsOnly="#orgContactsOnly#")>
			<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
				<CFOUTPUT><div id="error">#FORM.VerifyBy# does not have the permissions required in order to be listed as the #ClosureVerificationByLabel#.</div></CFOUTPUT>
				<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#PDA/PDASYNC.cfm" ACTION="Failure"></CFMODULE>
			</CFMODULE>
		</CFIF>

		<CFPARAM NAME="BlockCloseVerify" DEFAULT="true">
		<CFIF BlockCloseVerify IS true AND FORM.VerifyBy IS NOT "" AND FORM.VerifyBy IS Form.ClosePerson>
			<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
				<div id="error">The individual responsible for verification cannot be the same as the individual who closed the <cfoutput>#LCase(REQUEST.ATSFindingName)#</cfoutput>.</div>
				<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#PDA/PDASYNC.cfm" ACTION="Failure"></CFMODULE>
			</CFMODULE>
		</CFIF>

	</CFIF>
</CFIF>

<CFIF Form.Action Is "Add">
	<cfif RequireClosureVerificationOnAdd and FORM.VerifyBy IS "">
		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<div id="error">The individual responsible for verification cannot be blank.</div>
			<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#PDA/PDASYNC.cfm" ACTION="Failure"></CFMODULE>
		</CFMODULE>
	</cfif>

	<CFIF FindNoCase(">", ClosureList) EQ 0>
		<CFSET ClSpace = FindNoCase(" ", ClosureList)>
		<CFIF ClSpace GT 0>
			<CFSET ClDays = Trim(Left(ClosureList, ClSpace))>
			<CFIF IsNumeric(ClDays)>
				<CFSET ClosureDueDate = DateAdd("d", ClDays, Form.AuditDate)>
			<CFELSE>
				<CFSET ClosureDueDate = Form.ClosureDueDate>
			</CFIF>
		<CFELSE>
			<CFSET ClosureDueDate = Form.ClosureDueDate>
		</CFIF>
	<CFELSE>
		<cfset ClosureList = Replace(ClosureList, "> ", ">", "ONE")>
		<CFSET ClDays = Trim(Mid(ClosureList, FindNoCase(">", ClosureList)+1, FindNoCase(" ", ClosureList, FindNoCase(">", ClosureList))-FindNoCase(">", ClosureList)))>
		<CFIF Form.ClosureDueDate EQ "">
			<CFSET #ClosureDueDate# = #DateAdd("d", ClDays, Form.AuditDate)#>
		<CFELSE>
			<CFSET ClosureDueDate = Form.ClosureDueDate>
		</CFIF>
	</CFIF>
	<CFIF (#IsDate(ClosureDueDate)# Is NOT "Yes")>
		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<CFOUTPUT>
			<div id="error">#ClosureDueDateLabel# is Invalid.</div>
			<cfif Request.IsBlackBerry EQ True>
			<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
			Please use the Go Back button to return to your input form and re-submit.
			#thegobackform#
			<!--- <cfelse>
			#BACKLINK#
			 to go back and correct the input --->
			</cfif> .
			</CFOUTPUT>
		</CFMODULE>
	</CFIF>
		<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#datechk.cfm" DateValue="#ClosureDueDate#" ReturnBoolean="YES">
		<cfif PSValidDate EQ "NO" AND FORM.ClosureDueDate NEQ "">
			<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
				<cfoutput><div id="error">#ClosureDueDateLabel#, #ClosureDueDate#, is Invalid.</div> Please
				<cfif Request.IsBlackBerry EQ True>
				<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
				Please use the Go Back button to return to your input form and re-submit.
				#thegobackform#
				<!--- <cfelse>
				#BACKLINK#
				 to go back to the form and submit again after correcting the date. --->
				</cfif>
				</cfoutput>
			</CFMODULE>
		</cfif>
	<!--- since it was changed above, force the form scope to be the same --->
	<cfset FORM.ClosureDueDate = ClosureDueDate>
	<CFSET REDIRECTDATA.FORM.ClosureDueDate = Form.ClosureDueDate>

	<CFIF Len(Form.NumItems) GT 5>
		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<CFOUTPUT>
			<div id="error">The Number of Items should not exceed 5 characters!</div>
			<cfif Request.IsBlackBerry EQ True>
			<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
			Please use the Go Back button to return to your input form and re-submit.
			#thegobackform#
			<!--- <cfelse>
			#BACKLINK#
			 to go back and correct the input. --->
			</cfif>
			</CFOUTPUT>
		</CFMODULE>
	</CFIF>
	<CFIF #DateCompare(Form.AuditDate, ClosureDueDate)# GT 0>
		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<CFOUTPUT>
			<div id="error">The #ClosureDueDateLabel# for the #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# must be equal to or later than the #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# date!</div>
			<cfif Request.IsBlackBerry EQ True>
			Please use the Go Back button to return to your input form and re-submit.
			<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
			#thegobackform#
			<!--- <cfelse>
			#BACKLINK#
			 to go back and correct the input. --->
			</cfif>
			</CFOUTPUT>
		</CFMODULE>
	</CFIF>
	<cfset todaydate=DateFormat(now(),'dd-mmm-yyyy')>
	<cfset form.auditDate = DateFormat(form.auditDate,'dd-mmm-yyyy')>
	<cfif structKeyExists(form, 'currentTZ')&&isDate(form.currentTZ)>
		<cfset todaydate = form.currentTZ>
	</cfif>
    <CFIF #DateCompare(form.auditDate, todaydate)# GT 0>
        <CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
            <CFOUTPUT>
            <div id="error">The #REQUEST.ATSFindingName# Date for the #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# cannot be in the future!</div>
            <cfif Request.IsBlackBerry EQ True>
            Please use the Go Back button to return to your input form and re-submit.
            <!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
            #thegobackform#
            <!--- <cfelse>
            #BACKLINK#
             to go back and correct the input. --->
            </cfif>
            </CFOUTPUT>
        </CFMODULE>
    </CFIF>

	<!--- Added to stop BlackBerry users if these are't filled in. - CJ 10/18/2005 --->
	<!--- Added to stop BlackBerry users if these aren't filled in. - CJ 10/18/2005 --->
	<CFIF Form.Description IS "">
		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<CFOUTPUT>
			<div id="error">The #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# Description is required!</div>
			<cfif Request.IsBlackBerry EQ True>
			<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
			Please use the Go Back button to return to your input form and re-submit.
			#thegobackform#
			<!--- <cfelse>
			#BACKLINK#
			 to go back and correct the input. --->
			</cfif>
			</CFOUTPUT>
		</CFMODULE>
	</CFIF>

	<CFIF Form.CorrectiveAction IS "" AND RequireCorrectiveAction IS "yes">
		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<CFOUTPUT>
			<div id="error">The Corrective Action field is required!</div>
			<cfif Request.IsBlackBerry EQ True>
			<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
			Please use the Go Back button to return to your input form and re-submit.
			#thegobackform#
			<!--- <cfelse>
			#BACKLINK#
			 to go back and correct the input. --->
			</cfif>
			</CFOUTPUT>
		</CFMODULE>
	</CFIF>


	<CFQUERY Name="FindingNum" DATASOURCE="#ODBC#">
	<!--- SELECT Max(ID) AS MaxID FROM TblAudit
	GROUP BY Orgname, Location HAVING (Orgname=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#"> and Location='#Form.Location#') --->
	SELECT Max(ID) AS MaxID
	FROM TblAudit WITH (NOLOCK)
	WHERE (Orgname=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
		and Location=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.Location#">)
	</CFQUERY>
	<CFSET #NewID# = 1> <!--- Added for the case that there are no findings to begin with --->
<!--- 	<CFOUTPUT QUERY="FindingNum">
	<CFSET #NewID# = #MaxID# + 1>
	</CFOUTPUT> --->
	<!--- Changed to fix issue w/batch submits - Chuck Jody 05/09/2005 --->
	<CFIF FindingNum.MaxID IS NOT "">
		<CFSET NewID = FindingNum.MaxID + 1>
	</CFIF>

	<cfset History = #DateFormat(Now(),"mm/dd/yyyy")# & "," & #TimeFormat(Now(),"hh:mm:ss tt")# & "|">
	<cfif #SingleSignOn# eq "Yes"><cfset History = #History# & #AccessName#><cfelse><cfset History = #History# & "Unknown"></cfif>
	<cfset History = #History# & "|Create">
	<!--- ClassificationType, --->
	<cfquery name="DoesClosureListExist" DATASOURCE="#ODBC#">
		SELECT Closure
		FROM ltbClosure WITH (NOLOCK)
		WHERE Closure = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.ClosureList#">
	</cfquery>
	<cfif DoesClosureListExist.RecordCount EQ 0>
		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<CFOUTPUT>
				<div id="error">Sorry, the classification selected is not valid.</div>
				<!--- #BACKLINK# to go back and correct the input. --->
				</font>
			</CFOUTPUT>
		</CFMODULE>
	<cfelse>
		<!-- valid classification (closurelist) -->
	</cfif>
	<!--- check verifyby person  --->
	<cfif VerifyBy IS NOT "">
		<cfquery name="DoesVerifyPersonExist" datasource="#ODBC#">
			SELECT Contact_Name
			FROM ltbContact WITH (NOLOCK)
			WHERE Contact_Name = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#VerifyBy#">
		</cfquery>
		<cfif DoesVerifyPersonExist.RecordCount EQ 0>
			<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
				<CFOUTPUT>
					<div id="error">Sorry, the verify person chosen is not valid.</div>
					<!--- #BACKLINK# to go back and correct the input. --->
					</font>
				</CFOUTPUT>
			</CFMODULE>
		<cfelse>
			<!-- valid verify person -->
		</cfif>
	</cfif>

	<!--- check category --->
	<cfif Form.Category NEQ "">
		<cfquery name="TempQ" DATASOURCE="#ODBC#">
			SELECT TOP 1 Category
			FROM ltbCategory WITH (NOLOCK)
			WHERE Category = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.Category#">
		</cfquery>
		<cfif TempQ.RecordCount EQ 0>
			<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
				<CFOUTPUT>
					<div id="error">Sorry, the category chosen is not valid.</div>
					<!--- #BACKLINK# to go back and correct the input. --->
					</font>
				</CFOUTPUT>
			</CFMODULE>
		<cfelse>
			<!-- valid category -->
		</cfif>
	</cfif>

	<CFSET bRan = false>
	<CFSET iAttempt = 0>
	<CFLOOP CONDITION="#bRan# IS false">
		<CFSET bRan = true>
		<CFTRY>
		<CFSET sAccess = "Unknown">
		<cfif SingleSignOn eq "Yes">
			<CFSET sAccess = AccessName>
		</cfif>

		<!--- fetch schema for additional fields and add to struct so we know what max length the field is. --->
		<CFIF isDefined("additionalATSFields") AND isStruct(additionalATSFields)>
			<cfquery name="qAdditionalFieldLengths" datasource="#ODBC#">
				SELECT Column_Name, Character_Maximum_Length
				FROM information_schema.columns
				WHERE Table_Name = N'tblAudit'
					AND Data_Type = N'nvarchar'
			</cfquery>

			<CFLOOP LIST="#StructKeyList(additionalATSFields)#" INDEX="additionalGroup">
				<CFLOOP ARRAY="#additionalATSFields[additionalGroup]#" INDEX="Field">
					<cfset NewField = StructCopy(Field) />
					<cfset NewField.MaxLength = 0 />

					<cfloop query="qAdditionalFieldLengths">
						<cfif qAdditionalFieldLengths.Column_Name eq Field.Name>
							<cfset NewField.MaxLength = qAdditionalFieldLengths.Character_Maximum_Length />
							<cfbreak />
						</cfif>
					</cfloop>

					<cfset additionalATSFields[additionalGroup][ArrayFind(additionalATSFields[additionalGroup], Field)] = NewField />
				</CFLOOP>
			</CFLOOP>
		</CFIF>
		
		<CFQUERY Name="NewFinding" DATASOURCE="#ODBC#">
			INSERT INTO TblAudit
			(Orgname, Location,ID, AuditDate, AuditType, FindingType, Category, NumItems,
			RepeatItem, Classification, COE, Bldg, Workstation, AuditName, ResponPerson,
			ClosureDueDate, Citation, Description, CorrectiveAction, ContactPerson,
			ContactPhone,
			CLOSEDATE,
			CLOSECOMMENT,
			CLOSEPERSON,
			Status,
			RefType, RefID,
			UpdateDate, UpdateUser, UpdateHistory
			<!--- 08/16/02 --->
			,VerifyPerson
			<!--- 08/29/02 --->
			,DaysBeforeReminder
			<!--- 04/03/03 ---->
			,VerifyByDate
			<!--- 04/20/04 --->
			, SubCOEID
			, MULTICC
			<!--- 06/17/04 --->
			<CFIF ATS_CAPA_ENABLED IS true>
			, InvestigationDetails
			, #RootCauseColumn#
			, #EffectiveDateColumn#
			, #CAPARequiredColumn#
			, #EffectivityColumn1#
			, #EffectivityColumn2#
			, #RootCause1Column#
			, #RootCause2Column#
			</CFIF>
		<CFIF bHasSubCat>
			,#SubCategoryColumn#
		</CFIF>
		<CFIF RiskCategoryEnabled>
			, ClassificationType
		</CFIF>
			<!--- Auto Escalation fields --->
			<cfif structKeyExists(form, "AutoEscalation")>
				, AutoEscalation
				, EscalateDueDate
			</cfif>
			, ExternalSubmit
			<!--- additional fields --->
			<CFIF isDefined("additionalATSFields") AND isStruct(additionalATSFields)>
				<CFLOOP LIST="#StructKeyList(additionalATSFields)#" INDEX="additionalGroup">
					<CFLOOP ARRAY="#additionalATSFields[additionalGroup]#" INDEX="Field">
						<CFIF Field.Name IS NOT "" && (!structKeyExists(field, "extensionsData"))>
							,#Field.Name#
						</CFIF>
					</CFLOOP>
				</CFLOOP>
			</CFIF>
			<cfif trim(form.replication) eq true>
				,MXEncode
			</cfif>
			<cfif structKeyExists(form, "employee")>
				,employee
			</cfif>
			<cfif variables.showCoResponPerson and structKeyExists(form, "coResponPerson") and (form.coResponPerson neq "")>
				,coResponPerson
			</cfif>
				)
	     	VALUES (
			<CF_QUERYPARAM VALUE="#ONAME#" CFSQLTYPE="CF_SQL_VARCHAR">,
			<CF_QUERYPARAM VALUE="#Form.Location#" CFSQLTYPE="CF_SQL_VARCHAR">,
			<CF_QUERYPARAM VALUE="#NewID#" CFSQLTYPE="CF_SQL_BIGINT">,
			<CF_QUERYPARAM VALUE="#Form.AuditDate#" CFSQLTYPE="CF_SQL_TIMESTAMP">,
			<CF_QUERYPARAM VALUE="#Form.AuditType#" CFSQLTYPE="CF_SQL_VARCHAR">,
			<CF_QUERYPARAM VALUE="#Form.FindingType#" CFSQLTYPE="CF_SQL_VARCHAR">,
			<CF_QUERYPARAM VALUE="#Form.Category#" CFSQLTYPE="CF_SQL_VARCHAR">,
			<CF_QUERYPARAM VALUE="#VAL(Form.NumItems)#" CFSQLTYPE="CF_SQL_INTEGER">,
			<CF_QUERYPARAM VALUE="#VAL(Form.RepeatItem)#" CFSQLTYPE="CF_SQL_BIT">,
			<CF_QUERYPARAM VALUE="#Form.ClosureList#" CFSQLTYPE="CF_SQL_VARCHAR">,
			<CF_QUERYPARAM VALUE="#FORM.COE#" NULL="#isBlank(FORM.COE)#" CFSQLTYPE="CF_SQL_VARCHAR">,
			<CF_QUERYPARAM VALUE="#FORM.Bldg#" NULL="#isBlank(FORM.Bldg)#" CFSQLTYPE="CF_SQL_VARCHAR">,
			<CF_QUERYPARAM VALUE="#FORM.Workstation#" NULL="#isBlank(FORM.Workstation)#" CFSQLTYPE="CF_SQL_VARCHAR" MAXLENGTH="50">,
			<CF_QUERYPARAM VALUE="#FORM.AuditName#" NULL="#isBlank(FORM.AuditName)#" CFSQLTYPE="CF_SQL_VARCHAR" MAXLENGTH="100">,
			<CF_QUERYPARAM VALUE="#FORM.ResponPerson#" NULL="#isBlank(FORM.ResponPerson)#" CFSQLTYPE="CF_SQL_VARCHAR">,
			<CF_QUERYPARAM VALUE="#ClosureDueDate#" CFSQLTYPE="CF_SQL_TIMESTAMP">,
			<CF_QUERYPARAM VALUE="#FORM.CITATION#" NULL="#isBlank(FORM.CITATION)#" CFSQLTYPE="CF_SQL_VARCHAR" MAXLENGTH="#variables.citationMaxLength+5#">,
			<CF_QUERYPARAM VALUE="#FORM.Description#" NULL="#isBlank(FORM.Description)#" CFSQLTYPE="CF_SQL_LONGVARCHAR">,
			<CF_QUERYPARAM VALUE="#FORM.CorrectiveAction#" NULL="#isBlank(FORM.CorrectiveAction)#" CFSQLTYPE="CF_SQL_LONGVARCHAR">,
			<CF_QUERYPARAM VALUE="#FORM.ContactPerson#" NULL="#isBlank(FORM.ContactPerson)#" CFSQLTYPE="CF_SQL_VARCHAR">,
			<CF_QUERYPARAM VALUE="#FORM.ContactPhone#" NULL="#isBlank(FORM.ContactPhone)#" CFSQLTYPE="CF_SQL_VARCHAR" MAXLENGTH="50">,

			<cfif CLOSECOMMENT NEQ ""	AND	CLOSEDATE NEQ "" AND CLOSEPERSON NEQ "">
				<CF_QUERYPARAM VALUE="#FORM.CLOSEDATE#" CFSQLTYPE="CF_SQL_TIMESTAMP">,
				<CF_QUERYPARAM VALUE="#FORM.CLOSECOMMENT#" CFSQLTYPE="CF_SQL_LONGVARCHAR">,
				<CF_QUERYPARAM VALUE="#FORM.CLOSEPERSON#" CFSQLTYPE="CF_SQL_VARCHAR">,
				<CF_QUERYPARAM VALUE="Closed" CFSQLTYPE="CF_SQL_VARCHAR">,
			<cfelse>
				<CF_QUERYPARAM NULL="Yes" CFSQLTYPE="CF_SQL_DATE">,
				<CF_QUERYPARAM NULL="Yes" CFSQLTYPE="CF_SQL_LONGVARCHAR">,
				<CF_QUERYPARAM NULL="Yes" CFSQLTYPE="CF_SQL_VARCHAR">,
				<CF_QUERYPARAM VALUE="Open" CFSQLTYPE="CF_SQL_VARCHAR">,
			</cfif>
			<CF_QUERYPARAM VALUE="#Form.RefType#" NULL="#isBlank(FORM.RefType)#" CFSQLTYPE="CF_SQL_VARCHAR" MAXLENGTH="100">,
			<CF_QUERYPARAM VALUE="#Form.RefID#" NULL="#isBlank(FORM.RefID)#" CFSQLTYPE="CF_SQL_VARCHAR" MAXLENGTH="50">,

			<CF_QUERYPARAM VALUE="#now()#" CFSQLTYPE="CF_SQL_TIMESTAMP">,

			<CF_QUERYPARAM VALUE="#sAccess#" CFSQLTYPE="CF_SQL_VARCHAR">,
			<CF_QUERYPARAM VALUE="#History#" CFSQLTYPE="CF_SQL_LONGVARCHAR">,
			<CF_QUERYPARAM VALUE="#VerifyBy#" NULL="#isBlank(VerifyBy)#" CFSQLTYPE="CF_SQL_VARCHAR">,
			<CF_QUERYPARAM VALUE="#DaysBeforeReminder#" NULL="#isBlank(DaysBeforeReminder)#" CFSQLTYPE="CF_SQL_VARCHAR">,
			<CF_QUERYPARAM VALUE="#VerifyByDate#" NULL="#!isDate(VerifyByDate)#" CFSQLTYPE="CF_SQL_TIMESTAMP">,
			<CF_QUERYPARAM VALUE="#VAL(FORM.SubCOEID)#" NULL="#isBlank(FORM.SubCOEID)#" CFSQLTYPE="CF_SQL_BIGINT">,
			<CF_QUERYPARAM VALUE="#FORM.MULT_CC#" NULL="#isBlank(FORM.MULT_CC)#" CFSQLTYPE="CF_SQL_VARCHAR">,
			<!--- 06/17/04 --->
			<CFIF ATS_CAPA_ENABLED IS true>
				<CF_QUERYPARAM VALUE="#FORM.InvestigationDetails#" NULL="#isBlank(FORM.InvestigationDetails)#" CFSQLTYPE="CF_SQL_LONGVARCHAR">,
				<CF_QUERYPARAM VALUE="#Left(FORM.RootCause, 600)#" NULL="#isBlank(Left(FORM.RootCause, 600))#" CFSQLTYPE="CF_SQL_VARCHAR" MAXLENGTH="600">,
				<CF_QUERYPARAM VALUE="#DateFormat(FORM.EffectiveDate, "mm/dd/yyyy")#" NULL="#!isDate(FORM.EffectiveDate)#" CFSQLTYPE="CF_SQL_VARCHAR" BLANKNULL="true">,
				<CF_QUERYPARAM VALUE="#VAL(FORM.CAPARequired)#" CFSQLTYPE="CF_SQL_BIT">,

				<CFIF Len(FORM.Effectivity) GT 100>
					<CFSET Effectivity2 = Right(FORM.Effectivity, Len(FORM.Effectivity) - 100)>
				</CFIF>
				<CF_QUERYPARAM VALUE="#FORM.Effectivity#" NULL="#isBlank(Left(FORM.Effectivity, 100))#" CFSQLTYPE="CF_SQL_VARCHAR" MAXLENGTH="100" BLANKNULL="true">,

				<CF_QUERYPARAM VALUE="#Effectivity2#" NULL="#isBlank(Effectivity2)#" CFSQLTYPE="CF_SQL_VARCHAR" BLANKNULL="true">,
				<CF_QUERYPARAM VALUE="#FORM.RootCause1#" NULL="#isBlank(FORM.RootCause1)#" CFSQLTYPE="CF_SQL_VARCHAR">,
				<CF_QUERYPARAM VALUE="#FORM.RootCause2#" NULL="#isBlank(FORM.RootCause2)#" CFSQLTYPE="CF_SQL_VARCHAR">,
			</CFIF>
		<CFIF bHasSubCat>
			<CF_QUERYPARAM VALUE="#FORM.subCat#" NULL="#isBlank(FORM.subCat)#" CFSQLTYPE="CF_SQL_VARCHAR">,
		</CFIF>
		<CFIF RiskCategoryEnabled>
			<CF_QUERYPARAM VALUE="#FORM.RiskCategory#" NULL="#isBlank(FORM.RiskCategory)#" CFSQLTYPE="CF_SQL_BIGINT">,
		</CFIF>
		<!--- Auto Escalation fields --->
		<cfif structKeyExists(form, "AutoEscalation")>
			#VAL(form.AutoEscalation)#,
			<CF_QUERYPARAM VALUE="#ClosureDueDate#" CFSQLTYPE="CF_SQL_TIMESTAMP">,
		</cfif>
		<CF_QUERYPARAM VALUE="#VAL(Form.ExternalSubmit)#" CFSQLTYPE="CF_SQL_BIT">
		<!--- additional fields --->
		<CFIF isDefined("additionalATSFields") AND isStruct(additionalATSFields)>
			<CFLOOP LIST="#StructKeyList(additionalATSFields)#" INDEX="additionalGroup">
				<CFLOOP ARRAY="#additionalATSFields[additionalGroup]#" INDEX="Field">
					<CFIF Field.Name IS NOT "" && (!structKeyExists(field, "extensionsData"))>
						<CFPARAM NAME="FORM.#Field.Name#" DEFAULT="">
						,<CF_QUERYPARAM VALUE="#FORM[Field.Name]#" CFSQLTYPE="CF_SQL_VARCHAR" MAXLENGTH="#Field.MaxLength GT 0 ? Field.MaxLength : Len(FORM[Field.Name])#">
					</CFIF>
				</CFLOOP>
			</CFLOOP>
		</CFIF>
		<cfif trim(form.replication) eq true>
			,'0'
		</cfif>
		<cfif structKeyExists(form, "employee")>
			,<CF_QUERYPARAM VALUE="#FORM.employee#" NULL="#isBlank(FORM.employee)#" CFSQLTYPE="CF_SQL_VARCHAR">
		</cfif>
		<cfif variables.showCoResponPerson and structKeyExists(form, "coResponPerson") and (form.coResponPerson neq "")>
			,<CF_QUERYPARAM value="#FORM.coResponPerson#" NULL="#isBlank(FORM.coResponPerson)#" CFSQLTYPE="CF_SQL_VARCHAR">
		</cfif>
			)
		</CFQUERY>

		<CFCATCH TYPE="Any">
			<CFIF iAttempt GT 3>
				<CFRETHROW>
			<CFELSE>
				<CFSET bRan = false>
			</CFIF>
			<CFSET iAttempt = iAttempt + 1>
			<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#SLEEP.cfm" SECONDS="1">
			<CFQUERY Name="FindingNum" DATASOURCE="#ODBC#">
				SELECT Max(ID) AS MaxID
				FROM TblAudit WITH (NOLOCK)
				WHERE (Orgname=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
				and Location=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.Location#">)
			</CFQUERY>
			<CFSET NewID = 1>
			<CFIF FindingNum.MaxID IS NOT "">
				<CFSET NewID = FindingNum.MaxID + 1>
			</CFIF>
		</CFCATCH>
	</CFTRY>


	</CFLOOP>

	<CFSET Form.ID = NewID>
	<CFSET FORM.bAdded = true>

	<CFQUERY NAME="GetAuditID" DATASOURCE="#ODBC#">
		SELECT tblAuditID
		FROM tblAudit WITH (NOLOCK)
		WHERE ID = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_BIGINT" VALUE="#NewID#">
			AND Orgname = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
			AND Location = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.Location#">
	</CFQUERY>
	<CFSET FORM.tblAuditID = GetAuditID.tblAuditID>




	<!--- Log CR history for escalation to ATS : Loveena D 18th May 2016  --->
	<cfif Form.RefType NEQ "" AND Form.RefID NEQ "">
		<cfset escCFCObject = CreateObject("component", "#Library.CFC.DotPath#.escalationNotification").init(busDSN=ODBC)>
		<cfset escCFCObject.notifyApp(siteid=SiteID, sourceID=NewID, sourceType=request.ATSAppName, refID=Form.RefID, refType=Form.RefType, recordAction=Form.Action)>	
	</cfif>
	<cfif request.app.appid neq 3>
		<cfset request.app.appid = 3/>
	</cfif> 

	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#additionalFields.cfm"
					  siteid="#form.siteid#"
					  FIELDS="additionalATSFields" 
					  reftype="ats"
					  odbc="#odbc#"
					  orgid    = "#form.orgid#"
					  type="add"
					  refid="tblAuditID"
					  refidValue="#FORM.tblAuditID#"
					  datascope="#form#"
					  TRANSLATOR=""/>

	<!--- if approvals feature is active lets, process the approvals for them.  --->

	<cfset variables.queryExistingExtension = variables.extension.getRecords(scope_siteID=form.siteid,whereRecordName="tblauditId",whereRecordValue="#form.tblAuditId#",fieldlist="tblAuditId",dsn="#odbc#")/>
	<cfset variables['extensionRecordSet'] = {} />
	<cfloop collection="#form#" item = "element">
		<cfif findNoCase("approvals_", element)>
			<cfset variables['extensionRecordSet']["#element#"]	= form["#element#"] />
		</cfif>
	</cfloop>

	<cfif !structIsEmpty(variables['extensionRecordSet'])>
		<cfset variables['extensionRecordSet']["tblAuditId"] = form.tblAuditId />
		<cfif variables.queryExistingExtension.recordCount gt 0>
			<cfset variables.extension.updRecord(scope_siteID=form.siteid,record_group_id = variables.queryExistingExtension.record_group_id,update_by = request.user.accessName,recordset = variables.extensionRecordSet,insertIfNotExists = true,dsn=odbc)>
		<cfelse>
			<cfset variables.extension.insRecord(created_by=request.user.accessName,created_date = now(),scope_siteID=siteID,recordset =variables.extensionRecordSet)> 
		</cfif>	
	</cfif>

	
	<!--- ENDS approvals section on ADD mode --->
	
	<cfif variables.auditgpson>
		<!--- This creates the GPS record in the global GPS tables --->
		<cfif form.gpslong neq "" and form.gpslat neq "">
			<cfmodule
				template="#request.library.customtags.virtualpath#GPS/createGPS.cfm"
				odbc="#CC_ODBC#"
				refType="audfinding"
				refValue="#FORM.tblAuditID#"
				long="#FORM.GPSlong#"
				lat="#FORM.GPSlat#"
				comments="#FORM.GPScomments#"
			/>
		</cfif>
	</cfif>


	<!--- Cisneros	5/22/2006 --->
	<cfparam name="ECfile" default="0">
	<cfif isDefined("RadAttach")><cfset ECfile=1></cfif>

	<cfif CopyAttachment eq 1 and ECfile eq 1>
		<CFQUERY NAME="qGetDestinySiteIDAttach" DATASOURCE="#ODBC#" result="SQLSent">
			SELECT SiteID
			FROM Site WITH (NOLOCK)
			WHERE OrgName=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
			AND Location=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Loc#">
		</CFQUERY>

		<cfset NewIDAttach=NewID>

		<CFQUERY NAME="qGetSourceSiteIDAttach" DATASOURCE="#ODBC#">
			SELECT SiteID
			FROM Site WITH (NOLOCK)
			WHERE OrgName=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#CopyOrg#">
			 AND Location=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#CopyLoc#">
		</CFQUERY>

		<!--- BEGIN david zavalza 7/11/2006 --->
		<cfif listlen(RadAttach) NEQ 0>
			<cfset stInit = StructNew()>
			<cfset stInit.BusinessID = BusinessID>
			<cfset stInit.AppID = request.app.appID>
			<cfset stInit.SiteID = "#qGetSourceSiteIDAttach.SiteID#">
			<cfset stInit.ODBC = "#CC_ODBC#">
			<cfset objAttach = CreateObject("component", "#Library.CFC.DotPath#.attach").init(argumentCollection = stInit)>
			<cfset methodSuccess = objAttach.Copy(sourceAttachID=RadAttach,
												destinyRefID="#NewIDAttach#",
												destinyRefType="Audit",
												destinyBusinessID=BusinessID,
												destinyAppID=request.app.appID,
												destinySiteID="#qGetDestinySiteIDAttach.SiteID#")>
		</cfif>
		<!--- END david zavalza 7/11/2006 --->
	</cfif>	<!--- Copy attachment IF ends --->
</CFIF>

<!---Neha change begin for child action item: 237534--->
<CFSET CCDelList = "">
<CFIF isDefined("Form.SubCAUpdateStep")>
<CFLOOP LIST="#FORM.SubCAUpdateStep#" INDEX="iSubCACount">
	<CFQUERY NAME ="getAllNames" DATASOURCE="#ODBC#">
		SELECT Contact_Name, Contact_Email
			FROM ltbContact WITH (NOLOCK)
			WHERE Contact_Name =
			(SELECT Step_ResponPerson FROM TblAudit_Step WITH (NOLOCK)
			WHERE  OrgName = <CF_QUERYPARAM VALUE="#ONAME#" CFSQLTYPE="CF_SQL_VARCHAR"/>
				AND Location = <CF_QUERYPARAM VALUE="#Form.Location#" CFSQLTYPE="CF_SQL_VARCHAR"/>
				AND ID  = <CF_QUERYPARAM VALUE="#FORM.ID#" CFSQLTYPE="CF_SQL_BIGINT"/>
				AND Step_ID = <CF_QUERYPARAM VALUE="#iSubCACount#" CFSQLTYPE="CF_SQL_INTEGER"/>
			)
	</CFQUERY>
	<CFSET CCcountAll = 0>
	 <CFLOOP LIST="#CCList#" INDEX="icheck">
		 <CFIF ("CCList#icheck#" NEQ getAllNames.Contact_Email)>
		 	<CFSET CCcountAll = CCcountAll +1>
		 <CFELSE>
		 	<CFSET CCcountAll = -1>
			<CFBREAK>
		</CFIF>
	</CFLOOP>
	<CFSET CCcountAll = 0>
	 <CFLOOP LIST = "#CCDelList#" INDEX="icheck">
		 <CFIF ("CCDelList#icheck#" NEQ getAllNames.Contact_Email)>
		 	<CFSET CCcountAll = CCcountAll +1>
		 <CFELSE>
		 	<CFSET CCcountAll = -1>
			<CFBREAK>
		</CFIF>
	</CFLOOP>
	<CFIF CCcountAll NEQ -1>
		<CFSET CCList = ListAppend(CCList,getAllNames.Contact_Email)>
	</CFIF>
</CFLOOP>

<CFIF isDefined("FORM.SubCARemoveStep") and FORM.SubCARemoveStep IS NOT "">
	<CFLOOP LIST="#FORM.SubCARemoveStep#" INDEX="item">
		<CFIF ListFind(FORM.SubCAUpdateStep, item) NEQ 0>
			<CFQUERY NAME="qGetDelSub" DATASOURCE="#ODBC#">
				SELECT Step_ID, Step_Description, Step_ResponPerson, Step_ClosureDueDate, Step_Status, Step_CloseDate, Step_CloseComment
				FROM TblAudit_Step WITH (NOLOCK)
				WHERE OrgName = <CF_QUERYPARAM VALUE="#ONAME#" CFSQLTYPE="CF_SQL_VARCHAR"/>
					AND Location = <CF_QUERYPARAM VALUE="#Form.Location#" CFSQLTYPE="CF_SQL_VARCHAR"/>
					AND ID = <CF_QUERYPARAM VALUE="#FORM.ID#" CFSQLTYPE="CF_SQL_BIGINT"/>
					AND Step_ID = <CF_QUERYPARAM VALUE="#item#" CFSQLTYPE="CF_SQL_INTEGER"/>
			</CFQUERY>
			<CFSET FORM["RESPONPERSON#item#"] = #EncodeForHTML(qGetDelSub.Step_ResponPerson)#>
			<CFSET FORM["Desciption#item#"] = #EncodeForHTML(qGetDelSub.Step_Description)#>
			<CFSET FORM["ClosureDueDate#item#"] = #qGetDelSub.Step_ClosureDueDate#>
			<CFSET FORM["Status#item#"] = #EncodeForHTML(qGetDelSub.Step_Status)#>
			<CFSET FORM["CloseComment#item#"] = #qGetDelSub.Step_CloseDate#>
			<CFSET FORM["CloseDate#item#"] = #qGetDelSub.Step_CloseDate#>
			<CFQUERY NAME="qGetDelSubEmail" DATASOURCE="#ODBC#">
				SELECT Contact_Name, Contact_Email
				FROM ltbContact WITH (NOLOCK)
				WHERE Contact_Name = <CF_QUERYPARAM VALUE="#FORM["RESPONPERSON#item#"]#" CFSQLTYPE="CF_SQL_VARCHAR"/>
			</CFQUERY>
			<CFSET CCcountDel = 0>
			<CFIF FORM["RESPONPERSON#item#"] IS NOT "">
				 <CFLOOP LIST="#CCList#" INDEX="icheck">
					 <CFIF ("CCList#icheck#" NEQ qGetDelSubEmail.Contact_Email)>
					 	<CFSET CCcountDel = CCcountDel +1>
					 <CFELSE>
					 	<CFSET CCcountDel = -1>
						<CFBREAK>
					</CFIF>
				</CFLOOP>
			</CFIF>
			<CFIF CCcountDel NEQ -1>
				<CFSET CCDelList = ListAppend(CCDelList, qGetDelSubEmail.Contact_Email)>
			</CFIF>
		</CFIF>
	</CFLOOP>
</CFIF>
<!--- To check for updated RP --->

	 <CFLOOP LIST="#FORM.SubCAUpdateStep#" INDEX="iSubCACount">
		<CFPARAM NAME="FORM.ResponPerson#iSubCACount#" DEFAULT="">
		<CFPARAM NAME="FORM.ResponUpdPerson#iSubCACount#" DEFAULT="">
		<CFSET insertThisProperlyFormattedDate = "">
		<CFIF FORM["ResponPerson#iSubCACount#"] IS NOT "">

			<CFQUERY NAME ="UpdateNames" DATASOURCE="#ODBC#">
				SELECT Step_ID,Step_ResponPerson FROM TblAudit_Step WITH (NOLOCK)
				WHERE  OrgName = <CF_QUERYPARAM VALUE="#ONAME#" CFSQLTYPE="CF_SQL_VARCHAR"/>
					AND Location = <CF_QUERYPARAM VALUE="#Form.Location#" CFSQLTYPE="CF_SQL_VARCHAR"/>
					AND ID  = <CF_QUERYPARAM VALUE="#FORM.ID#" CFSQLTYPE="CF_SQL_BIGINT"/>
					AND Step_ID = <CF_QUERYPARAM VALUE="#iSubCACount#" CFSQLTYPE="CF_SQL_INTEGER"/>
					AND NOT EXISTS (
						SELECT *
						FROM tblAudit_Step WITH (NOLOCK)
						WHERE OrgName = <CF_QUERYPARAM VALUE="#ONAME#" CFSQLTYPE="CF_SQL_VARCHAR"/>
							AND Location = <CF_QUERYPARAM VALUE="#Form.Location#" CFSQLTYPE="CF_SQL_VARCHAR"/>
							AND ID  = <CF_QUERYPARAM VALUE="#FORM.ID#" CFSQLTYPE="CF_SQL_BIGINT"/>
							AND Step_ID = <CF_QUERYPARAM VALUE="#iSubCACount#" CFSQLTYPE="CF_SQL_INTEGER"/>
							AND Step_Description = <CF_QUERYPARAM VALUE="#FORM["Desciption#iSubCACount#"]#" CFSQLTYPE="CF_SQL_VARCHAR" MAXLENGTH="1500"/>
							AND Step_ResponPerson = <CF_QUERYPARAM VALUE="#FORM["ResponPerson#iSubCACount#"]#" CFSQLTYPE="CF_SQL_VARCHAR"/>
							AND Step_ClosureDueDate
							<CFIF FORM["ClosureDueDate#iSubCACount#"] IS "">
								IS NULL
							<CFELSE>
							<!--- 
								<cfset insertThisProperlyFormattedDate = DecodeForHTML(DecodeForHTML(DecodeForHTML(DecodeForHTML(DecodeForHTML(FORM["ClosureDueDate#iSubCACount#"])))))><!--- this better decode all of it --->
															<CFSET insertThisProperlyFormattedDate = DateFormat(LSParseDateTime(insertThisProperlyFormattedDate),'mm/dd/yyyy')> --->
							
								= <CF_QUERYPARAM VALUE="#FORM["CloseDate#iSubCACount#"]#" CFSQLTYPE="CF_SQL_DATE"/>
							</CFIF>
							AND Step_Status = <CF_QUERYPARAM VALUE="#FORM["Status#iSubCACount#"]#" CFSQLTYPE="CF_SQL_VARCHAR"/>
							AND Step_CloseDate
							<CFIF FORM["CloseDate#iSubCACount#"] IS "">
								IS NULL
							<CFELSE>
							<!--- 
								<cfset insertThisProperlyFormattedDate = DecodeForHTML(DecodeForHTML(DecodeForHTML(DecodeForHTML(DecodeForHTML(FORM["CloseDate#iSubCACount#"])))))>
															<cfset insertThisProperlyFormattedDate = DateFormat(LSParseDateTime(insertThisProperlyFormattedDate),'mm/dd/yyyy')> --->
							
								= <CF_QUERYPARAM VALUE="#FORM["CloseDate#iSubCACount#"]#" CFSQLTYPE="CF_SQL_DATE"/>
							</CFIF>
							AND Step_CloseComment
							<CFIF FORM["CloseComment#iSubCACount#"] IS "">
								IS NULL
							<CFELSE>
								= <CF_QUERYPARAM VALUE="#FORM["CloseComment#iSubCACount#"]#" CFSQLTYPE="CF_SQL_VARCHAR" MAXLENGTH="1500"/>
							</CFIF>
						)
			</CFQUERY>

			<CFIF #UpdateNames.Step_ID# IS NOT ''>
				<CFQUERY NAME="UpdateCheck" DATASOURCE="#ODBC#">
					SELECT Contact_Name, Contact_Email
					FROM ltbContact WITH (NOLOCK)
					WHERE Contact_Name = <CF_QUERYPARAM VALUE="#FORM["ResponPerson#UpdateNames.Step_ID#"]#" CFSQLTYPE="CF_SQL_VARCHAR"/>
				</CFQUERY>
				<CFSET FORM["ResponUpdPerson#iSubCACount#"] = #UpdateNames.Step_ResponPerson#>
				<CFQUERY NAME="BeforeUpdateCheck" DATASOURCE="#ODBC#">
					SELECT Contact_Name, Contact_Email
					FROM ltbContact WITH (NOLOCK)
					WHERE Contact_Name = <CF_QUERYPARAM VALUE="#FORM["ResponUpdPerson#iSubCACount#"]#" CFSQLTYPE="CF_SQL_VARCHAR"/>
				</CFQUERY>

				<!--- CC the main email to the SubCa user --->
				<CFSET CCcount1 = 0>
				<CFIF FORM["RESPONPERSON#iSubCACount#"] IS NOT "">
					 <CFLOOP LIST="#CCList#" INDEX="icheck">
						 <CFIF ("CCList#icheck#" NEQ UpdateCheck.Contact_Email)>
						 	<CFSET CCcount1 = CCcount1 +1>
						 <CFELSE>
						 	<CFSET CCcount1 = -1>
							<CFBREAK>
						</CFIF>
					</CFLOOP>
					<CFSET CCcount1 = 0>
					<CFIF FORM["RESPONPERSON#iSubCACount#"] IS NOT "">
						<CFLOOP LIST="#CCList#" INDEX="icheck">
							 <CFIF ("CCList#icheck#" NEQ BeforeUpdateCheck.Contact_Email)>
						 		<CFSET CCcount1 = CCcount1 +1>
							 <CFELSE>
							 	<CFSET CCcount1 = -1>
								<CFBREAK>
							</CFIF>
						</CFLOOP>
						<CFSET CCcount1 = 0>
						 <CFLOOP LIST = "#CCDelList#" INDEX="icheck">
							 <CFIF ("CCDelList#icheck#" NEQ UpdateCheck.Contact_Email)>
							 	<CFSET CCcount1 = CCcount1 +1>
							 <CFELSE>
							 	<CFSET CCcount1 = -1>
								<CFBREAK>
							</CFIF>
						</CFLOOP>
						<CFIF CCcount1 NEQ -1>
							<CFSET CCDelList = ListAppend(CCDelList, UpdateCheck.Contact_Email)>
						</CFIF>

						<CFSET CCcount1 = 0>
						  <CFLOOP LIST = "#CCDelList#" INDEX="icheck">
							 <CFIF ("CCDelList#icheck#" NEQ BeforeUpdateCheck.Contact_Email)>
							 	<CFSET CCcount1 = CCcount1 +1>
							 <CFELSE>
							 	<CFSET CCcount1 = -1>
								<CFBREAK>
							</CFIF>
						</CFLOOP>
						<CFIF CCcount1 NEQ -1>
							<CFSET CCDelList = ListAppend(CCDelList, BeforeUpdateCheck.Contact_Email)>
						</CFIF>
					</CFIF>
				</CFIF>
			</CFIF>
		</CFIF>
	</CFLOOP>
	</cfif>

	<CFSET FORM.CCDeleteList=CCDelList>

<!---Neha change end for Child Action Item: 237534--->
<!--- only passed by audit.cfc because I didn;t want to have to set the variables --->
<CFPARAM NAME="FORM.IgnoreSubCA" DEFAULT="false">
<CFIF FORM.Action IS NOT "View">
	<CFIF FORM.IgnoreSubCA IS false>
		<CFINCLUDE TEMPLATE="audaction_SubAct.cfm">
	</CFIF>
</CFIF>

<!--- 07/14/03 --->
<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#PDA/PDASync.cfm" ACTION="Success"></CFMODULE>


<CFIF #Form.Action# Is "Delete">
		
	<CFQUERY NAME="qSiteIDAttach" DATASOURCE="#ODBC#">
		SELECT SiteID, tblAuditID
		FROM Site WITH (NOLOCK)
			INNER JOIN tblAudit WITH (NOLOCK)
				ON SITE.OrgName = tblAudit.Orgname
				AND SITE.Location = tblAudit.Location
				AND tblAudit.ID = #FORM.ID#
		WHERE SITE.OrgName=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
			AND SITE.Location=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#FORM.Location#">
	</CFQUERY>
	<CFIF qSiteIDAttach.RecordCount NEQ 0>
		<!--- delete any attachments --->
		<cfset stInit = StructNew()>
		<cfset stInit.BusinessID = BusinessID>
		<cfset stInit.AppID = request.app.appID>
		<cfset stInit.SiteID = qSiteIDAttach.SiteID>
		<cfset stInit.ODBC = CC_ODBC>
		<cfset objAttach = CreateObject("component", "#Library.CFC.DotPath#.attach").init(argumentCollection = stInit)>
		<cfset qGetAttachments = objAttach.get(RefID=Form.ID)>
		<cfif qGetAttachments.RecordCount>
			<cfset methodSuccess = objAttach.Delete(AttachID=valuelist(qGetAttachments.AttachID))>
		</cfif>
	</CFIF>
	
	<cfquery name="variables.qTblAuditId" datasource="#odbc#">
		select tblAuditId
		from   tblAudit
		WHERE (Orgname=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
		and Location=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.Location#">
		and (ID=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_BIGINT" VALUE="#Form.ID#">));	
	</cfquery>
	<cfset variables.tblauditid = variables.qTblAuditId.tblauditid />

	<CFQUERY DATASOURCE="#ODBC#">
	Delete TblAudit
	WHERE (Orgname=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
		and Location=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.Location#">
		and (ID=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_BIGINT" VALUE="#Form.ID#">));
	</CFQUERY>
	<!--- 06/22/04 --->
	<CFQUERY DATASOURCE="#ODBC#">
		DELETE TblAudit_Archive
		WHERE Orgname=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
			AND Location=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.Location#">
			AND ID=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_BIGINT" VALUE="#Form.ID#">
	</CFQUERY>

	<!--- mb: remove any watch list items --->
	<cfif isnumeric(qSiteIDAttach.tblauditid)>
		<cfmodule template="#request.library.customTags.virtualPath#gsoTDwatch.cfm"
			refType = "Audit" <!--- do not change; internal use only, safest setting is a hardcoded value --->
			refID = "#qSiteIDAttach.tblauditid#"
			watchaction = "delete"
		/>
	</cfif>

</CFIF>

<!--- Copy attachments also --->
<CFQUERY NAME="qGetSiteID" DATASOURCE="#ODBC#">
	SELECT SiteID
	FROM Site WITH (NOLOCK)
	WHERE OrgName=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
	 AND Location=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Loc#">
</CFQUERY>
<CFSET LocID = qGetSiteID.SiteID>

<CFIF Form.Action IS "Add">


	<!--- handle attachments on add --->
	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#GetAttachments.cfm"
		Org="#Org#"
		Loc="#Loc#"
		AppName="#request.ATSAppName#"
		AttachParent="Audit"
        UpdateRefID="true"
		NewRefID="#FORM.ID#"
		SiteID="#LocID#">

	<CFIF FiveWhyLabel IS NOT "">
		<!--- 5Why on add --->
		<cfmodule template="#request.library.customTags.virtualPath#5Why.cfm" reftype="ats" refid="#FORM.tblAuditID#" reconcile="true">
	</CFIF>

</CFIF>

<!--- 06/17/04 --->
<CFIF FORM.Action IS "Edit" AND ATS_AUDIT_TRAIL IS true>
	<CFQUERY NAME="qGetFields" DATASOURCE="#ODBC#">
		SELECT *
		FROM TblAudit_Archive WITH (NOLOCK)
		WHERE OrgName =<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
			AND Location = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.Location#">
			AND ID = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_BIGINT" VALUE="#FORM.ID#">
	</CFQUERY>
	<CFSET ColumnList = qGetFields.columnList>
	<CFSET ColumnList = ListDeleteAt(ColumnList, ListFindNoCase(ColumnList, "VersionNo"))>

	<CFQUERY NAME="qGetHistory" DATASOURCE="#ODBC#">
		SELECT UpdateHistory
		FROM tblAudit  WITH (NOLOCK)
		WHERE Orgname=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
			AND Location = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.Location#">
			AND ID = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_BIGINT" VALUE="#Form.ID#">
	</CFQUERY>
	<CFSET sUpdateHistory = Replace(qGetHistory.UpdateHistory, "||", "|Unknown|", "ALL")>
	<CFSET iMissingHistory = ListLen(sUpdateHistory, "|")/3>
	<CFIF iMissingHistory GT qGetFields.RecordCount>
		<CFSET VersionNo = iMissingHistory>
	<CFELSE>
		<CFSET VersionNo = qGetFields.RecordCount + 1>
	</CFIF>

	<CFQUERY NAME="qFixHistory" DATASOURCE="#ODBC#">
		DELETE FROM TblAudit_Archive
		WHERE OrgName =<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
			AND Location = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.Location#">
			AND ID = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_BIGINT" VALUE="#FORM.ID#">
			AND VersionNo = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_BIGINT" VALUE="#VersionNo#">
	</CFQUERY>

	
	<CFQUERY NAME="qInsertHistory" DATASOURCE="#ODBC#">
		IF NOT EXISTS (
			SELECT VersionNo,#ColumnList#
			FROM TblAudit_Archive WITH (NOLOCK)
			WHERE OrgName =<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
				AND Location = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.Location#">
				AND ID = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_BIGINT" VALUE="#FORM.ID#">
				AND VersionNo = '#VersionNo#')
		BEGIN
			INSERT INTO TblAudit_Archive
			(VersionNo,#ColumnList#)
			SELECT #VersionNo#,#ColumnList#
			FROM TblAudit WITH (NOLOCK)
			WHERE OrgName =<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
				AND Location = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.Location#">
				AND ID = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_BIGINT" VALUE="#FORM.ID#">
		END
	</CFQUERY>

	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#GetAttachments.cfm"
		query="qGetAttachments"
		Org="#Org#"
		Loc="#Loc#"
		AppName="#request.ATSAppName#"
		AttachParent="Audit"
		AttachID="#FORM.ID#"
		BaseAttachPath="#AuditHome#attachments"
		SiteID="#LocID#" >

	<!--- BEGIN david zavalza 7/11/2006 --->
	<cfif qGetAttachments.RecordCount>
		<cfset stInit = StructNew()>
		<cfset stInit.BusinessID = BusinessID>
		<cfset stInit.AppID = request.app.appID>
		<cfset stInit.SiteID = "#LocID#">
		<cfset stInit.ODBC = "#ODBC#">
		<cfset objAttach = CreateObject("component", "#Library.CFC.DotPath#.attach").init(argumentCollection = stInit)>
		<cfset methodSuccess = objAttach.Copy(sourceAttachID=valuelist(qGetAttachments.AttachID),
											destinyRefID="#FORM.ID#.#VersionNo#",
											destinyRefType="Audit",
											destinyBusinessID=BusinessID,
											destinyAppID=request.app.appID,
											destinySiteID="#LocID#")>
	</cfif>
	<!--- END david zavalza 7/11/2006 --->

	<CFIF FiveWhyLabel IS NOT "">
		<!--- 5Why on replicate--->
		<cfmodule template="#request.library.customTags.virtualPath#5Why.cfm" reftype="ats" refid="#FORM.tblAuditID#" DuplicateRefID="#FORM.tblAuditID#.#VersionNo#">
	</CFIF>

</CFIF>

<CFIF FORM.Action IS NOT "Delete">

	<CFIF isDefined("Form.ClosureDueDate") AND Form.ClosureDueDate IS NOT "" AND NOT isDate(Form.ClosureDueDate)>
		<CFIF FindNoCase(">", ClosureList) EQ 0 AND RiskCategory_EditClosureDueDate eq "false">
			<CFSET ClSpace = FindNoCase(" ", ClosureList)>
			<CFIF ClSpace GT 0>
				<CFSET ClDays = Trim(Left(ClosureList, ClSpace))>
				<CFIF IsNumeric(ClDays)>
					<CFSET ClosureDueDate = DateAdd("d", ClDays, Form.AuditDate)>
				<CFELSE>
					<CFSET ClosureDueDate = Form.ClosureDueDate>
				</CFIF>
			<CFELSE>
				<CFSET ClosureDueDate = Form.ClosureDueDate>
			</CFIF>
		<CFELSE>
			<CFSET ClDays = Trim(Mid(ClosureList, FindNoCase(">", ClosureList)+1, FindNoCase(" ", ClosureList, FindNoCase(">", ClosureList))-FindNoCase(">", ClosureList)))>
			<CFIF Form.ClosureDueDate EQ "">
				<CFSET ClosureDueDate = #DateAdd("d", ClDays, Form.AuditDate)#>
			<CFELSE>
				<CFSET ClosureDueDate = Form.ClosureDueDate>
			</CFIF>
		</CFIF>

		<CFSET Form.ClosureDueDate = ClosureDueDate>
	</CFIF>

	<CFSET Details = ArrayNew(1)>
	<cfif findingDateLabel is not "">
        <CFSET DetailsSummary = ArrayNew(1)>
        <CFSET Details[ArrayLen(Details)+1] = StructNew()>
        <cfset Details[ArrayLen(Details)].Title = findingDateLabel>
        <cfset Details[ArrayLen(Details)].Value = DateFormat(Form.AuditDate, 'd-mmm-yy')>
    </cfif>
    <cfif findingTypeLabel is not "">
        <CFSET Details[ArrayLen(Details)+1] = StructNew()>
        <cfset Details[ArrayLen(Details)].Title = findingTypeLabel>
        <CFIF HighPriority EQ 1>
            <cfset Details[ArrayLen(Details)].Value = "<B>" & EncodeForHTML(Form.FindingType) & "<B>">
        <CFELSE>
            <cfset Details[ArrayLen(Details)].Value = EncodeForHTML(Form.FindingType)>
        </CFIF>
	</cfif>

	<cfif AuditActionTypeLabel is not "">
		<CFSET Details[ArrayLen(Details)+1] = StructNew()>
		<cfset Details[ArrayLen(Details)].Title = AuditActionTypeLabel>
		<cfset Details[ArrayLen(Details)].Value = EncodeForHTML(Form.AuditType)>
	</cfif>
    <CFIF Form.RiskCategory IS NOT "" and RiskCategoryLabel is not "">
		<CFQUERY NAME="qGetRiskCategory" DATASOURCE="#ODBC#">
			SELECT Words
			FROM ltbEnvPriority WITH (NOLOCK)
			WHERE ltbEnvPriority.EnvPriorityID = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_BIGINT" VALUE="#VAL(Form.RiskCategory)#">
		</CFQUERY>

		<CFSET Details[ArrayLen(Details)+1] = StructNew()>
		<cfset Details[ArrayLen(Details)].Title = RiskCategoryLabel>
		<cfset Details[ArrayLen(Details)].Value = EncodeForHTML(qGetRiskCategory.Words)>
	</CFIF>

	<cfif numberOfItemsLabel is not "">
		<CFSET Details[ArrayLen(Details)+1] = StructNew()>
		<cfset Details[ArrayLen(Details)].Title = numberOfItemsLabel>
		<cfset Details[ArrayLen(Details)].Value = EncodeForHTML(Form.NumItems)>
	</cfif>
	<cfif repeatFindingLabel is not "">
		<CFSET Details[ArrayLen(Details)+1] = StructNew()>
		<cfset Details[ArrayLen(Details)].Title = repeatFindingLabel>
		<cfset Details[ArrayLen(Details)].Value = EncodeForHTML(YesNoFormat(Form.RepeatItem))>
	</cfif>
	<cfif auditNameNumberLabel is not "">
		<CFSET Details[ArrayLen(Details)+1] = StructNew()>
		<cfset Details[ArrayLen(Details)].Title = auditNameNumberLabel>
		<CFIF Form.AuditName Is Not "">
			<cfset Details[ArrayLen(Details)].Value = EncodeForHTML(Form.AuditName)>
		<CFELSE>
			<cfset Details[ArrayLen(Details)].Value = "Not Specified">
		</CFIF>
	</cfif>
	<cfif findingCategoryLabel is not "">
		<CFSET Details[ArrayLen(Details)+1] = StructNew()>
		<cfset Details[ArrayLen(Details)].Title = findingCategoryLabel>
		<cfset Details[ArrayLen(Details)].Value = EncodeForHTML(Form.Category)>
	</cfif>
	<CFIF findingSubCategoryLabel is not "" and Form.subCat IS NOT "">
		<CFSET Details[ArrayLen(Details)+1] = StructNew()>
		<CFSET Details[ArrayLen(Details)].Title = findingSubCategoryLabel>
		<CFSET Details[ArrayLen(Details)].Value = EncodeForHTML(Form.subCat)>
	</CFIF>
	<!--- additional fields ---->
	<cfset AFExclude = "Finding Highlights,Closure Priority">
	<cfif structKeyExists(variables,"additionalATSFields") && isStruct(additionalATSFields) && !structIsEmpty(additionalATSFields)>
		<CFLOOP LIST="#StructKeyList(additionalATSFields)#" INDEX="additionalGroup">
			<CFLOOP ARRAY="#additionalATSFields[additionalGroup]#" INDEX="Field">
				<cfif structKeyExists(Field,"showOnConfirmation") && isValid("boolean",Field.showOnConfirmation) && Field.showOnConfirmation && additionalGroup is not "Closure Priority">
					<cfset AFExclude = listDeleteAt(AFExclude, ListFindNoCase(AFExclude,additionalGroup))>
					<cfbreak>
				</cfif>
			</cfloop>
		</cfloop>
	</cfif>
	
	<CFSET FieldData = StructNew()>
	<CFSET FieldLabels = StructNew()>
	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#additionalFields.cfm" 
		FIELDS="additionalATSFields" 
	    EXCLUDEGROUPS="#AFExclude#" 
	    TYPE="VARIABLES" 
	    reftype="ats"
	    orgid    = "#form.orgid#"
				  odbc="#odbc#"
				  refid="tblAuditID"
				  refidValue="#FORM.tblAuditID#" 
	    SCOPE="#FieldData#" 
	    LABELSCOPE="#FieldLabels#" 
	    DATASCOPE="FORM" />

	    <CFSET sharedRows = 0>
		<CFLOOP COLLECTION="#FieldData#" ITEM="FieldName">
			<CFIF structKeyExists(fieldLabels, fieldName) && FieldLabels[FieldName] IS NOT "">
				<CFIF !ListFindNoCase(additionalSharedRows,Left(FieldName,Len(FieldName)-3))>
					<CFSET Details[ArrayLen(Details)+1] = StructNew()>
					<cfset Details[ArrayLen(Details)].Title = FieldLabels[FieldName]>
					<cfset Details[ArrayLen(Details)].Value = EncodeForHTML(FieldData[FieldName])>
				<CFELSE>
					<CFSET sharedRows++>
				</CFIF>
			</CFIF>
		</CFLOOP>
		<cfif Len(additionalSharedRows) && sharedRows gt 0>
			<cfset sharedRows = sharedRows/ListLen(additionalSharedRows)>
			<cfloop from="1" to="#sharedRows#" index="r">
				<cfloop list="#additionalSharedRows#" index="shared">
					<cfif structKeyExists(fieldLabels, shared & numberFormat(r,'000'))>
						<CFSET Details[ArrayLen(Details)+1] = StructNew()>
						<cfset Details[ArrayLen(Details)].Title = Translator.Translate(FieldLabels[shared & numberFormat(r,'000')])>
						<cfset Details[ArrayLen(Details)].Value = EncodeForHTML(FieldData[shared & numberFormat(r,'000')])>
					</cfif>
				</cfloop>
			</cfloop>
		</cfif>

	<cfif responsiblePersonLabel is not "">
		<CFSET Details[ArrayLen(Details)+1] = StructNew()>
		<CFSET iResponsiblePersonLocation = ArrayLen(Details)>
		<cfset Details[ArrayLen(Details)].Title = responsiblePersonLabel>
		<cfset Details[ArrayLen(Details)].Value = Form.ResponPerson>
		<CFIF Form.ResponPerson IS NOT Form.ClosePerson AND Form.ClosePerson IS NOT "">
			<cfset Details[ArrayLen(Details)].Value = Form.ResponPerson & "<BR> &nbsp; &nbsp; <I>#REQUEST.ATSFindingName# closed by: #Form.ClosePerson#</i>">
		</CFIF>
	</cfif>
	<cfif variables.showCoResponPerson is true and (form.coResponPerson neq "")>
		<CFSET Details[ArrayLen(Details)+1] = StructNew()>
		<cfset Details[ArrayLen(Details)].Title = coResponPersonLabel>
		<cfset Details[ArrayLen(Details)].Value = Form.coResponPerson>
	</cfif>
	<cfif variables.autoEscalation_enable>
		<!--- Autoescalation: Active/Inactive --->
		<cfif structKeyExists(Form, 'AutoEscalation')>
			<cfset Details[ArrayLen(Details)+1] = StructNew()>
			<cfset Details[ArrayLen(Details)].Title = Translator.Translate("Auto Escalation Enabled")>
			<cfif Form.AutoEscalation EQ 0>
				<cfset Details[ArrayLen(Details)].Value = Translator.Translate("No")>
			<cfelse>
				<cfset Details[ArrayLen(Details)].Value = Translator.Translate("Yes")>
			</cfif>
		</cfif>
		<!--- Autoescalation: Escalated Responsible Person --->
		<cfif structKeyExists(Form, 'escalatedResponPerson')>
			<cfset Details[ArrayLen(Details)+1] = StructNew()>
			<cfset Details[ArrayLen(Details)].Title = Translator.Translate("Auto Escalation Responsible Person")>
			<cfset Details[ArrayLen(Details)].Value = Form.escalatedResponPerson>
		</cfif>
		<!--- Autoescalation: Due Date --->
		<cfif structKeyExists(form, 'EscalateDueDate') AND structKeyExists(Form, 'escalatedResponPerson')>
			<cfset Details[ArrayLen(Details)+1] = StructNew()>
			<cfset Details[ArrayLen(Details)].Title = Translator.Translate("Auto Escalation Due Date")>
			<cfset Details[ArrayLen(Details)].Value = DateFormat(form.EscalateDueDate, "dd-mmm-yyyy")>
		</cfif>
	</cfif>
	<CFIF isDefined("Form.CloseDate") AND Form.CloseDate IS NOT "">
		<CFSET Details[ArrayLen(Details)+1] = StructNew()>
		<cfset Details[ArrayLen(Details)].Title = "Date #LCase(REQUEST.ATSFindingName)# was closed">
		<cfset Details[ArrayLen(Details)].Value = DateFormat(Form.CloseDate, "dd-mmm-yyyy")>
	</CFIF>

<cfif CenterDeptLabel is not "">
	<CFSET Details[ArrayLen(Details)+1] = StructNew()>
	<cfset Details[ArrayLen(Details)].Title = CenterDeptLabel>

	<CFIF Form.COE Is Not "">
		<cfset Details[ArrayLen(Details)].Value = EncodeForHTML(Form.COE)>
		<CFIF SubCOEName IS NOT "">
			<cfset Details[ArrayLen(Details)].Value = EncodeForHTML(Form.COE) & "/" & EncodeForHTML(SubCOEName)>
		</CFIF>
	<CFELSE>
		<cfset Details[ArrayLen(Details)].Value = "Not Specified">
	</CFIF>
</cfif>
<cfif BuildingLabel is not "">
	<CFSET Details[ArrayLen(Details)+1] = StructNew()>
	<cfset Details[ArrayLen(Details)].Title = BuildingLabel>
	<CFIF Form.Bldg Is Not "">
		<cfset Details[ArrayLen(Details)].Value = EncodeForHTML(Form.Bldg)>
	<CFELSE>
		<cfset Details[ArrayLen(Details)].Value = "Not Specified">
	</CFIF>
</cfif>
<cfif WorkstationLabel is not "">
	<CFSET Details[ArrayLen(Details)+1] = StructNew()>
	<cfset Details[ArrayLen(Details)].Title = WorkstationLabel>
	<CFIF Form.Workstation Is Not "">
		<cfset Details[ArrayLen(Details)].Value = EncodeForHTML(Form.Workstation)>
	<CFELSE>
		<cfset Details[ArrayLen(Details)].Value = "Not Specified">
	</CFIF>
</cfif>
<cfif AuditorLabel is not "">
	<CFSET Details[ArrayLen(Details)+1] = StructNew()>
	<cfset Details[ArrayLen(Details)].Title = AuditorLabel>
	<cfset Details[ArrayLen(Details)].Value = Form.ContactPerson>
</cfif>
	<CFIF multipleEmailLabel is not "" and FORM.MULT_CC IS NOT "">
		<CFSET Details[ArrayLen(Details)+1] = StructNew()>
		<cfset Details[ArrayLen(Details)].Title = multipleEmailLabel>
		<cfset Details[ArrayLen(Details)].Value = Replace(Replace(FORM.MULT_CC, ",", ", ", "All"), ";", "; ", "All")>
	</CFIF>
<cfif FindingDescriptionLabel is not "">
	<CFSET DetailsSummary[ArrayLen(DetailsSummary)+1] = StructNew()>
	<cfset DetailsSummary[ArrayLen(DetailsSummary)].Title = FindingDescriptionLabel>
	<!--- <cfset DetailsSummary[ArrayLen(DetailsSummary)].Value = audit.TextFormat(string=Form.Description,encodeMethod="HTML")> --->
	<cfset DetailsSummary[ArrayLen(DetailsSummary)].Value = Form.Description>
</cfif>

<cfif CitationLabel is not "">
	<CFSET Details[ArrayLen(Details)+1] = StructNew()>
	<cfset Details[ArrayLen(Details)].Title = CitationLabel>
	<cfset Details[ArrayLen(Details)].Value = EncodeForHTML(Form.Citation)>
</cfif>

	<CFIF ATS_CAPA_ENABLED IS true>
		<cfif subjectCAPALabel is not "">
            <CFSET Details[ArrayLen(Details)+1] = StructNew()>
            <cfset Details[ArrayLen(Details)].Title = subjectCAPALabel>
            <cfset Details[ArrayLen(Details)].Value = "Yes">
        </cfif>
        <cfif InvestigationDetailsLabel is not "">
			<CFSET Details[ArrayLen(Details)+1] = StructNew()>
			<cfset Details[ArrayLen(Details)].Title = InvestigationDetailsLabel>
			<cfset Details[ArrayLen(Details)].Value = EncodeForHTML(FORM.InvestigationDetails)>
		</cfif>
		<cfif RootCauseLabel is not "">
			<CFSET Details[ArrayLen(Details)+1] = StructNew()>
			<cfset Details[ArrayLen(Details)].Title = RootCauseLabel>
			<cfset Details[ArrayLen(Details)].Value = EncodeForHTML(FORM.RootCause)>
		</cfif>
		<cfif bRootCauseDropDowns eq true && structKeyExists(variables, "JSONRootCause")>
			<CFSET Details[ArrayLen(Details)+1] = StructNew()>
			<cfset Details[ArrayLen(Details)].Title = RootCause1Label>
			<cfset Details[ArrayLen(Details)].Value = EncodeForHTML(FORM.RootCause1)>

			<CFSET Details[ArrayLen(Details)+1] = StructNew()>
			<cfset Details[ArrayLen(Details)].Title = RootCause2Label>
			<cfset Details[ArrayLen(Details)].Value = EncodeForHTML(FORM.RootCause2)>
		</cfif>
		<cfif EffectivityDateLabel is not "">
			<CFSET Details[ArrayLen(Details)+1] = StructNew()>
			<cfset Details[ArrayLen(Details)].Title = EffectivityDateLabel>
			<CFIF FORM.EffectiveDate IS NOT "">
				<CFSET FORM.EffectiveDate = DateFOrmat(FORM.EffectiveDate, "dd-mmm-yyyy")>
			</CFIF>
			<cfset Details[ArrayLen(Details)].Value = FORM.EffectiveDate>
		</cfif>
		<cfif EffectivityInformationLabel is not "">
			<CFSET Details[ArrayLen(Details)+1] = StructNew()>
			<cfset Details[ArrayLen(Details)].Title = EffectivityInformationLabel>
			<cfset Details[ArrayLen(Details)].Value = EncodeForHTML(FORM.Effectivity)>
		</cfif>
	</CFIF>

	<cfif CorrectiveActionLabel is not "">
		<CFSET DetailsSummary[ArrayLen(DetailsSummary)+1] = StructNew()>
		<cfset DetailsSummary[ArrayLen(DetailsSummary)].Title = CorrectiveActionLabel>
		<cfset DetailsSummary[ArrayLen(DetailsSummary)].Value = Form.CorrectiveAction>
	</cfif>

	<CFIF ClosureCommentLabel is not "" and isDefined("Form.CloseComment") AND Form.CloseComment IS NOT "">
		<CFSET Details[ArrayLen(Details)+1] = StructNew()>
		<cfset Details[ArrayLen(Details)].Title = ClosureCommentLabel>
		<cfset Details[ArrayLen(Details)].Value = audit.TextFormat(string=Form.CloseComment,encodeMethod="HTML")>
	</CFIF>

	<cfif ClosureDueDateLabel is not "" and IsDefined("Form.ClosureDueDate") AND Form.ClosureDueDate IS NOT "" AND IsDate(Form.ClosureDueDate)>
		<CFSET DetailsSummary[ArrayLen(DetailsSummary)+1] = StructNew()>
		<cfset DetailsSummary[ArrayLen(DetailsSummary)].Title = ClosureDueDateLabel>
		<cfset DetailsSummary[ArrayLen(DetailsSummary)].Value = DateFormat(Form.ClosureDueDate, 'd-mmm-yy')>
	</cfif>

	<!--- additional fields  --->
	<CFSET FieldData = StructNew()>
	<CFSET FieldLabels = StructNew()>
	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#additionalFields.cfm" FIELDS="additionalATSFields"
     INCLUDEGROUPS="Closure Priority" TYPE="VARIABLES" SCOPE="#FieldData#" LABELSCOPE="#FieldLabels#" DATASCOPE="FORM" orgid    = "#form.orgid#" />

    <CFLOOP COLLECTION="#FieldData#" ITEM="FieldName">
		<CFIF structKeyExists(FieldLabels, FieldName) && FieldLabels[FieldName] IS NOT "">
			<CFSET DetailsSummary[ArrayLen(DetailsSummary)+1] = StructNew()>
			<cfset DetailsSummary[ArrayLen(DetailsSummary)].Title = FieldLabels[FieldName]>
			<cfset DetailsSummary[ArrayLen(DetailsSummary)].Value = EncodeForHTML(FieldData[FieldName])>
		</CFIF>
	</CFLOOP>

	<CFINCLUDE TEMPLATE="audaction_subMail.cfm">


	<CFIF FORM.VerifyBy IS NOT "" and ClosureVerificationByLabel is not "">
		<CFSET Details[ArrayLen(Details)+1] = StructNew()>
		<CFSET iVerifyByLocation = ArrayLen(Details)>
		<cfset Details[ArrayLen(Details)].Title = ClosureVerificationByLabel>
		<cfset Details[ArrayLen(Details)].Value = FORM.VerifyBy>
		<CFIF VerifyByDate IS NOT "" and ScheduledVerificationDateLabel is not "">
			<CFSET Details[ArrayLen(Details)+1] = StructNew()>
			<cfset Details[ArrayLen(Details)].Title = ScheduledVerificationDateLabel>
			<cfset Details[ArrayLen(Details)].Value = DateFormat(VerifyByDate, 'd-mmm-yy')>
		</CFIF>
	</CFIF>

	<CFIF VerifyComment IS NOT "">
		<CFIF VerifyDate IS NOT "" and ClosureVerificationDateLabel is not "">
			<CFSET Details[ArrayLen(Details)+1] = StructNew()>
			<cfset Details[ArrayLen(Details)].Title = ClosureVerificationDateLabel>
			<cfset Details[ArrayLen(Details)].Value = DateFormat(VerifyDate, 'd-mmm-yy')>
		</CFIF>
		<cfif ClosureVerificationCommentLabel is not "">
			<CFSET Details[ArrayLen(Details)+1] = StructNew()>
			<cfset Details[ArrayLen(Details)].Title = ClosureVerificationCommentLabel>
			<cfset Details[ArrayLen(Details)].Value = audit.TextFormat(string=VerifyComment,encodeMethod="HTML")>
		</cfif>
	</CFIF>

	 <CFIF variables.auditgpson and GPSLabel is not "">
        <cfsavecontent variable="GPSDisplay">
            <cfmodule 
                template="#request.library.customtags.virtualpath#GPS/displayGPS.cfm"
                odbc="#CC_ODBC#"
                refType="audfinding"
                refValue="#FORM.tblAuditID#"
            />
        </cfsavecontent>

        <CFSET Details[ArrayLen(Details)+1] = StructNew()>
        <cfset Details[ArrayLen(Details)].Title = GPSLabel>
        <cfset Details[ArrayLen(Details)].Value = GPSDisplay>
    </CFIF>

	<!--- add attachments to email --->
	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#GetAttachments.cfm"
		query="qGetAttachments"
		Org="#Org#"
		Loc="#Loc#"
		AppName="#request.ATSAppName#"
		AttachParent="Audit"
		AttachID="#ID#"
		BaseAttachPath="#AuditHome#attachments"
		SiteID="#LocID#" >
	<STYLE>
		.noimage img
		{
			display:none;
		}
	</STYLE>
	<CFSAVECONTENT VARIABLE="sAttach">
			<cfloop query="qGetAttachments">
				<CFIF qGetAttachments.FolderID IS NOT "">
					<cfoutput>
					<img src="#request.domainprotocol##request.domainurl##request.library.images.url#icons/16x16/icon_folder_16x16.gif" BORDER="0">
						#qGetAttachments.displayname#<BR>
					</cfoutput>
				<CFELSE>
					<CFSET sAttachLink = qGetAttachments.weblink>
					<CFIF sAttachLink IS "">
						<CFSET sAttachLink = "#request.domainprotocol##request.domainurl##qGetAttachments.filepath#">
					</CFIF>
					<CFIF qGetAttachments.PARENTFOLDERID IS NOT "0">
						&nbsp;&nbsp;
					</CFIF>
					<cfoutput>
					<img src="#request.domainprotocol##request.domainurl##request.library.images.url#bullets/blue_bullet.gif" alt="" border="0" align="absmiddle"
					 style="margin: 0px 5px 0px 5px;">
					<a href="#sAttachLink#">#qGetAttachments.displayname#</a><br />
					</cfoutput>
				</CFIF>
			</cfloop>
	</CFSAVECONTENT>
	<!--- Add the attachments into the structure --->
	<cfif fileAttachmentsLabel is not "">
        <CFSET Details[ArrayLen(Details)+1] = StructNew()>
        <CFSET Details[ArrayLen(Details)].Title = fileAttachmentsLabel>
        <CFSET Details[ArrayLen(Details)].Value = sAttach>
    </cfif>
	
	<cfset extensionDetails = "">
	<cfmodule template="#request.library.customtags.virtualpath#extensionshandler.cfm"
			appid="#request.app.appID#" ref_id="#tblAuditID#" action="emailDetails"
			audithome="#audithome#" odbc="#odbc#" siteid = "#locid#" returnVar="extensionDetails" />
	<cfif isArray(extensionDetails)>
		<cfset details.addAll(extensionDetails) />
		<cfloop array="#details#" index="field">
		 	<cfif trim(field.title) eq "Long Term Verifier">
		 		<cfset ltv_verifier = field.value>
		 	</cfif>
		</cfloop>
	</cfif>
	
</CFIF>



<CFSAVECONTENT VARIABLE="Footer">
	<CFOUTPUT>
		<cfset coRPLinkFooter=""/>
		<cfset coRPFooter=""/>
		<cfif variables.showCoResponPerson and (form.coResponPerson neq "")>
			<cfset coRPLinkFooter="&CoRP=#EncodeForURL(Form.coResponPerson)#"/>
			<cfset coRPFooter=" or #coResponPersonLabel#"/>
		</cfif>
	<div ALIGN="right" class="mediumtext"><A target="_top" HREF="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#audit.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&RP=#EncodeForURL(Form.ResponPerson)##coRPLinkFooter#&Crit=Yes"><b>Click here</b></A> to view ALL open #LCase(REQUEST.ATSAuditName)# #LCase(REQUEST.ATSFindingName)#s at this Site for which you are designated as the Responsible Person#coRPFooter#!</div>
	</CFOUTPUT>
</CFSAVECONTENT>

<cfset Form.Status = ListFirst(Form.Status, ",")>


<CFSET StatusText = audit.GetStatusTextForDisplay(Status="#Form.Status#", CILINK="#CI_actions_comp#", VerifyDate="#VerifyDate#", VERIFYPERSON="#VerifyBy#", ReOpenStatus="#ReOpenEmail#", Action="#FORM.Action#", EmailStatusText="True")>


<!--- 08/16/02 --->
<CFIF FORM.Action IS "Edit" OR FORM.Action IS "Reject">

	<CFIF FORM.Action IS "Reject">
		<CFSET FORM.Status = "Open">
		<CFSET FORM.ClosePerson = "">
		<CFSET FORM.CloseDate = "">
	</CFIF>

	<cfset ClosureList = Replace(ClosureList,"]",">")>
	<CFIF FindNoCase(">", ClosureList) EQ 0 AND RiskCategory_EditClosureDueDate eq "false">
		<CFSET ClSpace = FindNoCase(" ", ClosureList)>
		<CFIF ClSpace GT 0>
			<CFSET ClDays = Trim(Left(ClosureList, ClSpace))>
			<!--- Used to be IsNumeric(ClDays) --->
			<CFIF val(ClDays) GT 0 and NOT (DateFormat(Form.ClosureDueDate,'dd-mmm-yyyy') NEQ DateFormat(DateAdd("d", ClDays, Form.AuditDate),'dd-mmm-yyyy') AND  isDate(Form.ClosureDueDate))>
				<CFSET ClosureDueDate = DateAdd("d", ClDays, Form.AuditDate)>
			<CFELSE>
				<CFSET ClosureDueDate = Form.ClosureDueDate>
			</CFIF>
		<CFELSE>
			<CFSET ClosureDueDate = Form.ClosureDueDate>
		</CFIF>
	<CFELSE>
		<CFSET ClDays = Trim(Mid(ClosureList, FindNoCase(">", ClosureList)+1, FindNoCase(" ", ClosureList, FindNoCase(">", ClosureList))-FindNoCase(">", ClosureList)))>
		<CFIF Form.ClosureDueDate EQ "">
			<CFSET ClosureDueDate = #DateAdd("d", ClDays, Form.AuditDate)#>
		<CFELSE>
			<CFSET ClosureDueDate = Form.ClosureDueDate>
		</CFIF>
	</CFIF>
	<CFIF (#IsDate(ClosureDueDate)# Is NOT "Yes")>
		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<CFOUTPUT>
			<div id="error">The #ClosureDueDateLabel# for the #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# is not a valid date.</div>
			<!--- #BACKLINK# to go back and correct the input. --->
			</CFOUTPUT>
		</CFMODULE>
	</CFIF>

	<CFIF #DateCompare(Form.AuditDate, ClosureDueDate)# GT 0>
		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<CFOUTPUT>
			<div id="error">The #ClosureDueDateLabel# for the #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# must be equal to or later than the #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# date!</div>
			<!--- #BACKLINK# to go back and correct the input. --->
			</CFOUTPUT>
		</CFMODULE>
	</CFIF>
	<cfset todaydate=DateFormat(now(),'dd-mmm-yyyy')>
	<cfif structKeyExists(form, 'currentTZ') AND Form.currentTZ NEQ "">
		<cfset todaydate = form.currentTZ>
	</cfif>
    <CFIF #DateCompare(Form.AuditDate, todaydate)# GT 0>
        <CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
            <CFOUTPUT>
            <div id="error">The #REQUEST.ATSFindingName# Date for the #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# cannot be in the future!</div>
            <cfif Request.IsBlackBerry EQ True>
            Please use the Go Back button to return to your input form and re-submit.
            <!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
            #thegobackform#
            <!--- <cfelse>
            #BACKLINK#
             to go back and correct the input. --->
            </cfif>
            </CFOUTPUT>
        </CFMODULE>
    </CFIF>
	
	<!--- 05/28/03 --->
	<CFQUERY NAME="qGetDate" DATASOURCE="#ODBC#">
		SELECT VerifyByDate, VerifyPerson, MultiCC, tblAuditID
		FROM tblAudit WITH (NOLOCK)
		WHERE OrgName =<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
			AND Location = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.Location#">
			AND ID = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_BIGINT" VALUE="#Form.ID#">
	</CFQUERY>
	<cfset FORM.tblAuditID = qGetDate.tblAuditID>
	<CFIF qGetDate.VerifyByDate IS NOT "" AND
	 DateFormat(qGetDate.VerifyByDate, "mm/dd/yyyy") IS NOT DateFormat(FORM.VerifyByDate, "mm/dd/yyyy")>
		<!--- 07/24/03 --->
		<CFIF FORM.VerifyBy IS "">
			<CFSET sVerifyBy = qGetDate.VerifyPerson>
		<CFELSE>
			<CFSET sVerifyBy = FORM.VerifyBy>
		</CFIF>

		<CFQUERY NAME="qGetEmail" DATASOURCE="#ODBC#">
			SELECT c1.Contact_Email AS RPEmail, c2.Contact_Email AS VEmail
			FROM ltbContact c1 WITH (NOLOCK), ltbContact c2 WITH (NOLOCK)
			WHERE c1.Contact_Name = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#FORM.ResponPerson#">
				AND c2.Contact_Name = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#sVerifyBy#">
		</CFQUERY>
		<cfif findingTypeLabel eq "">
			<cfset SubjectLine = "#Abr#  #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# ##" & FORM.ID & " #ScheduledVerificationDateLabel# modified: " & OName & ", " & EncodeForHTML(FactoryID) & Loc & " (#GEBusiness#)">
		<cfelse>
			<CFSET SubjectLine = "#EncodeForHTML(Form.FindingType)# #Abr#  #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# ##" & FORM.ID & " #ScheduledVerificationDateLabel# modified: " & OName & ", " & EncodeForHTML(FactoryID) & Loc & " (#GEBusiness#)">
		</cfif>

	<CFIF qGetEmail.VEmail IS NOT "">
		<!--- 04/20/04 --->
		<CFIF isDefined("Form.EmailRP") AND Form.EmailRP IS NOT "">
			<CFSET CCList = ListAppend(qGetEmail.RPEmail, qGetDate.MultiCC)>
		</CFIF>
		<cfif variables.showCoResponPerson and (form.coResponPerson neq "")>
			<CFQUERY NAME="qGetCoRPEmail" DATASOURCE="#ODBC#">
				SELECT Contact_Email AS CoRPEmail
				FROM ltbContact WITH (NOLOCK)
				WHERE Contact_Name = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#FORM.coResponPerson#">
			</CFQUERY>
			<CFSET CCList = ListAppend(CCList, qGetCoRPEmail.CoRPEmail)>
		</cfif>
		<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#CLEANEMAIL.cfm"
			ToList = "#qGetEmail.VEmail#"
			CCList = "#CCList#"
			Delim = ",">

		<!--- "[FindingType] Type Finding Notification" distribution lists setup in contacts permissions if activated for this business   --->
		<cfmodule template="#ContactsMasterPath#distGetList.cfm"
			Application = "#request.ATSAppName#"
		    List = "#Form.FindingType# Type Finding Notification"
			AdditionalToEmails="#ToList#"
			AdditionalCcEmails="#CCList#"
		    SiteID ="#LocID#"
			DSN="#ODBC#">
		<cfset ToList = DistList.ToEmails>
		<cfset CCList = DistList.ccEmails>

		<CFSAVECONTENT VARIABLE="EmailNotes">
			<CFOUTPUT>
			<!--- 11/06/02 --->
				#AccessName# has updated this #LCase(REQUEST.ATSFindingName)# and changed the scheduled
				Closure Verification Date
				from <B>#DateFormat(qGetDate.VerifyByDate, "dd-mmm-yyyy")#</B> to
				 <FONT COLOR="blue">
				<CFIF FORM.VerifyByDate IS "">
					<STRONG>verification not required!</STRONG></div>
				<CFELSE>
				 	<B>#DateFormat(FORM.VerifyByDate, "dd-mmm-yyyy")#</B>!</div>
				</CFIF>
				 </FONT>
			</CFOUTPUT>
		</CFSAVECONTENT>

		<CFIF FORM.Status IS "Open">
			<CFSET StatusText =  "Open">
		</CFIF>
		<!--- this email does't show the verifyby row since it is to the verifyby --->
		<CFIF isDefined("iVerifyByLocation")>
			<CFSET ArrayDeleteAt(Details, iVerifyByLocation)>
			<CFSET iVerifyByLocation = -1>
		</CFIF>

		<!--- Closure Verification Date change email
		http://cincep09corpge.corporate.ge.com/help/ccenter/ats/ATS_Verification_Date_Changed.mht --->
		<CFMODULE TEMPLATE="#REQUEST.LIBRARY.CUSTOMTAGS.VIRTUALPATH#EmailDetails.cfm"
		 	Details="#Details#"
			DetailsTitle="#REQUEST.ATSFindingName# Details"
		 	Summary="#DetailsSummary#"
			SummaryTitle="#REQUEST.ATSFindingName# Summary"
			ApplicationName="#request.ATSAppName#"
			ApplicationIcon="ats.gif"
			ApplicationURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#audit.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&FactoryID=#EncodeForURL(factoryid)#"
			OrganizationName="#OName#"
			SiteName="#factoryid##Form.Location#"
			BlockTitle="#REQUEST.ATSFindingName# ID##"
			BlockID="#Form.ID#"
			BlockIDURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#status.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&ID=#EncodeForURL(FORM.ID)#&FactoryID=#EncodeForURL(factoryid)#"
			EmailType="#REQUEST.ATSAuditName# #REQUEST.ATSFindingName# Summary"
			EmailNotes="#EmailNotes#"
			StatusTitle="Status"
			StatusText="#StatusText#"
			LINKURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#audfinding.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&ID=#EncodeForURL(FORM.ID)#&FactoryID=#EncodeForURL(factoryid)#"
			LINKTEXT="Edit this #REQUEST.ATSAuditName# #REQUEST.ATSFindingName#"
			 TO="#ToList#" 
			 CC="#CCList#" 
		<!--- 	to="diego.serrano@gensuitellc.com" --->
			FROM="#EMailSender#"
			SUBJECT="#SubjectLine#"
			MailParam="X-Priority:1,X-MSMail-Priority:High"
			>

	</CFIF>
	</CFIF>

	<cfset GetHistory.UpdateHistory = "">
	<cfquery name="GetHistory" DATASOURCE="#ODBC#">
		select UpdateHistory, Status
		from tblAudit  WITH (NOLOCK)
		where Orgname=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
		and Location = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.Location#">
		and  ID = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_BIGINT" VALUE="#Form.ID#">
	</cfquery>
	<cfif #GetHistory.UpdateHistory# neq "">
		<cfset History = #GetHistory.UpdateHistory# & "|" & #DateFormat(Now(),"mm/dd/yyyy")# & "," & #TimeFormat(Now(),"hh:mm:ss tt")# & "|">
	<cfelse>
		<cfset History = #DateFormat(Now(),"mm/dd/yyyy")# & "," & #TimeFormat(Now(),"hh:mm:ss tt")# & "|">
	</cfif>
	<cfif #SingleSignOn# eq "Yes"><cfset History = #History# & #AccessName#><cfelse><cfset History = #History# & "Unknown"></cfif>
	<!--- 08/16/02 --->
	<cfset UpdateVerifyBy = 0>
	<CFIF FORM.Action IS "Reject">
		<CFSET History = History & "|Rejected">
	<CFELSEIF isDate(VerifyDate) AND FORM.Status IS "Closed">
		<CFSET History = History & "|Closure Verified">
	<CFELSEIF #GetHistory.Status# eq #Form.Status#>
		<cfset History = #History# & "|Edit">
	<cfelse>
		<!--- 6/3/04 --->
		<cfif VerifyBy NEQ ""><!--- If old VerifyBy does not exist in ltbContact, set to AccessName --->
			<cfquery name="quickcheck" datasource="#ODBC#">
				SELECT 1
				FROM ltbContact WITH (NOLOCK)
				WHERE Contact_Name=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#VerifyBy#">
			</cfquery>
			<cfif quickcheck.RecordCount EQ 0><cfset UpdateVerifyBy = 1><cfset VerifyBy = AccessName></cfif>
		</cfif>
		<cfif #Form.Status# eq "Open"><BR>
			<cfset History = #History# & "|Re-Open">
			<!--- 09/15/02 --->
			<!--- reset the verification information --->
			<CFSET VerifyDate = "">

			<!--- 11/06/02 --->
			<CFSET ReOpenEmail = "yes">
		</cfif>
		<cfif #Form.Status# eq "Closed"><cfset History = #History# & "|Close"></cfif>
	</cfif>

	<!--- 09/20/04 --->
	<CFTRY>

	<!--- check verifyby person ONLY if still open and it has changed OR they would now have some responsibilities--->
	<cfif VerifyBy NEQ "" AND (VerifyBy IS NOT VerifyBy_Orig OR VerifyDate IS "")>
		<cfquery name="DoesVerifyPersonExist" DATASOURCE="#ODBC#">
			SELECT Contact_Name
			FROM ltbContact WITH (NOLOCK)
			WHERE Contact_Name = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#VerifyBy#">
		</cfquery>
		<cfif DoesVerifyPersonExist.RecordCount EQ 0>
			<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
				<CFOUTPUT>
				<div id="error">
					Sorry, the verify person chosen is not valid.
					<!--- #BACKLINK# to go back and correct the input. --->
				</div>
				</CFOUTPUT>
			</CFMODULE>

		<cfelse>
			<!-- valid verify person -->
		</cfif>
	</cfif>
	<!--- Ensure that if the this is open that the close date/person are blank --->
	<CFIF FORM.Status IS "Open">
		<CFSET Form.CloseDate = "">
		<CFSET Form.ClosePerson = "">
	</CFIF>

	<cfquery name="qTemp" datasource="#odbc#">
		SELECT ClosureDueDate,Status
		FROM TblAudit WITH (NOLOCK)
		WHERE OrgName =<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
			AND Location = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.Location#">
			AND ID = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_BIGINT" VALUE="#FORM.ID#">
	</cfquery>
	<cfif SendEmailWhenFindingClosureDateModified neq "NO" and qTemp.RecordCount>
		<!--- CHECK PREVIOUS CLOSURE DUE DATE SO WE CAN COMPARE TO THE NEW ONE AND SEND EMAIL IF THEY NEED EMAIL ON CLOSURE DUE DATE CHANGE --->
		<cfset PreviousClosureDueDate = ParseDateTime(qTemp.ClosureDueDate)>
		<cfset NewClosureDueDate = ParseDateTime(ClosureDueDate)>

		<!--- IF IT IS TRUE then send an email to indicate CLOSUREDUEDATE was changed --->
		<CFIF  PreviousClosureDueDate neq NewClosureDueDate>
			<CFSET ToList = "">
			<CFSET CCList = "">
			<CFQUERY NAME="qGetResponsibleEmail" DATASOURCE="#ODBC#">
				SELECT CONTACT_EMAIL
				FROM ltbContact WITH (NOLOCK)
				WHERE Contact_Name = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.ResponPerson#">
			</CFQUERY>
			<CFIF qGetResponsibleEmail.CONTACT_EMAIL IS "">
				<CFQUERY NAME="qGetResponsibleEmail" DATASOURCE="#ODBC#">
					SELECT CONTACT_EMAIL
					FROM ltbContact WITH (NOLOCK)
					WHERE Contact_Name = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#AccessName#">
				</CFQUERY>
			</CFIF>
			<CFSET ToList = ListAppend(ToList, qGetResponsibleEmail.CONTACT_EMAIL)>
			<cfset FromEmail = qGetResponsibleEmail.CONTACT_EMAIL>
			<cfif variables.showCoResponPerson and (form.coResponPerson neq "")>
				<CFQUERY NAME="qGetCoResponsibleEmail" DATASOURCE="#ODBC#">
					SELECT CONTACT_EMAIL
					FROM ltbContact WITH (NOLOCK)
					WHERE Contact_Name = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.coResponPerson#">
				</CFQUERY>
				<CFSET ToList = ListAppend(ToList, qGetCoResponsibleEmail.CONTACT_EMAIL)>
			</cfif>
			<CFQUERY NAME="qGetResponsibleEmailCC" DATASOURCE="#ODBC#">
				SELECT CONTACT_EMAIL
				FROM ltbContact WITH (NOLOCK)
				WHERE Contact_Name = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.ContactPerson#">
			</CFQUERY>
			<CFSET CCList = ListAppend(CCList, qGetResponsibleEmailCC.CONTACT_EMAIL)>
			<CFIF FORM.MULT_CC IS NOT "">
				<CFSET CCList = ListAppend(CCList, FORM.MULT_CC)>
			</CFIF>
			<CFIF ToList IS NOT "">
				<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#CLEANEMAIL.cfm"
					ToList = "#ToList#"
					CCList = "#CCList#"
					Delim = ",">
				<CFSAVECONTENT VARIABLE="EmailNotes">
					<CFOUTPUT>
						<b>Dear <font color="blue"><i>#FORM.ResponPerson#:</b></i></font>
						<br>
						The #ClosureDueDateLabel# of the #LCase(REQUEST.ATSAuditName)# #LCase(REQUEST.ATSFindingName)# below has been changed from #DateFormat(PreviousClosureDueDate, "dd-mmm-yy")# to #DateFormat(NewClosureDueDate, "dd-mmm-yy")#<cfif IsDefined("AccessName")> by #AccessName#</cfif>.
					</CFOUTPUT>
				</CFSAVECONTENT>
				<cfif findingTypeLabel eq "">
					<cfset SubjectLine = "#Abr#  #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# ##" & FORM.ID & " #ClosureDueDateLabel# Modified: " & OName & ", " & EncodeForHTML(FactoryID) & Loc & " (#GEBusiness#)">
				<cfelse>
					<CFSET SubjectLine = "#EncodeForHTML(Form.FindingType)# #Abr#  #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# ##" & FORM.ID & " #ClosureDueDateLabel# Modified: " & OName & ", " & EncodeForHTML(FactoryID) & Loc & " (#GEBusiness#)">
				</cfif>

				<CFSET inFile = "">
				<cfset FileName = "ATS_Finding_#ID#_" & RandRange(0,9999)>

				<!--- Write File: Task Closure form to send back to Task Owner after All SubTasks are Closed- CJ 05/03/2007 --->
				<CFIF FORM.ExternalSubmit EQ 1>
					<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#WriteOfflinePage.cfm"
						Action="WriteFile"
						PageURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#status.cfm"
						FileName="#FileName#"
						SiteID="#LocID#"
						ID="#ID#"
						ExternalSubmitForm="true"
						AttachedFileName="inFile">

					<CFIF inFile IS NOT "">
						<CFSAVECONTENT VARIABLE="Footer">
							<CFOUTPUT>#Footer#</CFOUTPUT>
							<!--- External Submit --->
							<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#ExternalSubmit.cfm"
									Action="ExternalSubmitInfo"/>
						</CFSAVECONTENT>
					</CFIF>
				</CFIF>

				<!--- "[FindingType] Type Finding Notification" distribution lists setup in contacts permissions if activated for this business   --->
				<cfmodule template="#ContactsMasterPath#distGetList.cfm"
					Application = "#request.ATSAppName#"
				    List = "#Form.FindingType# Type Finding Notification"
					AdditionalToEmails="#ToList#"
					AdditionalCcEmails="#CCList#"
				    SiteID ="#LocID#"
					DSN="#ODBC#">
				<cfset ToList = DistList.ToEmails>
				<cfset CCList = DistList.ccEmails>

				<!--- Closure Due Date Email
				http://cincep09corpge.corporate.ge.com/help/ccenter/ats/Closure_Due_Date_Updated.mht --->
				<CFMODULE TEMPLATE="#REQUEST.LIBRARY.CUSTOMTAGS.VIRTUALPATH#EmailDetails.cfm"
					Details="#Details#"
					DetailsTitle="#REQUEST.ATSFindingName# Details"
					Summary="#DetailsSummary#"
					SummaryTitle="#REQUEST.ATSFindingName# Summary"
					ApplicationName="#request.ATSAppName#"
					ApplicationIcon="ats.gif"
					ApplicationURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#audit.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&FactoryID=#EncodeForURL(factoryid)#"
					OrganizationName="#OName#"
					SiteName="#factoryid##Loc#"
					BlockTitle="#REQUEST.ATSFindingName# ID##"
					BlockID="#ID#"
					BlockIDURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#status.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&ID=#ID#&FactoryID=#EncodeForURL(factoryid)#"
					EmailType="#REQUEST.ATSAuditName# #REQUEST.ATSFindingName# Summary"
					EmailNotes="#EmailNotes#"
					StatusTitle="Status" StatusText="#qtemp.status#" StatusTextColor="blue"
					LINKURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#audfinding.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&ID=#ID#&FactoryID=#EncodeForURL(factoryid)#"
					LINKTEXT="Edit this #REQUEST.ATSAuditName# #REQUEST.ATSFindingName#"
					 TO="#ToList#" 
					 CC="#CCList#" 
					<!--- to="diego.serrano@gensuitellc.com" --->
					FROM="#FromEmail#"
					SUBJECT="#SubjectLine#"
					FOOTER="#Footer#"
					MimeAttach="#inFile#"
					>
			<CFSAVECONTENT VARIABLE="sAdditionalMessage">
				<p>
				<!--- inform user --->
				<font class="extramediumtext"><i>
				<CFIF ToList EQ "">
				The designated Responsible Person does not have a valid email address, therefore no email was sent.
				<CFELSE>
					<CFOUTPUT>
					An email on this #LCase(REQUEST.ATSAuditName)# #LCase(REQUEST.ATSFindingName)# has been sent to inform the responsible person at <b>#ToList#</b>
					<CFIF CCList IS NOT "">
						(with a cc to <b>#CCList#</b>)
					</CFIF>
					stating the #ClosureDueDateLabel# was modified!<p>
					</CFOUTPUT>
				</CFIF>
				</i></font>
			</CFSAVECONTENT>
				<cfset EmailAlreadySent = 1>
			</CFIF>
		</CFIF>
	</cfif>

	<!--- temp till I figure out that is going on --->
	<CFIF isDate(VerifyByDate) AND Year(VerifyByDate) EQ 2000>
		<CFSET VerifyByDate = DateFormat(VerifyByDate, "mm/dd/") & "2009">

		<CFMAIL TO="ActionTrackingSystem.PM@gensuitellc.com" FROM="ActionTrackingSystem.PM@gensuitellc.com" SUBJECT="VerifyByDate BAD" TYPE="HTML">
			<CFDUMP VAR="#CGI#">
			<CFDUMP VAR="#FORM#">
		</CFMAIL>
	</CFIF>


		<CFSET Citation = Replace(Form.Citation, "|", " ", "ALL")>
		<CFSET Description = Replace(Form.Description, "|", " ", "ALL")>
		<CFSET CorrectiveAction = Replace(Form.CorrectiveAction, "|", " ", "ALL")>
		<CFSET sAccessName = "Unknown">
		<cfif SingleSignOn eq "Yes">
			<CFSET sAccessName = AccessName>
		</cfif>
		<CFSET Effectivity2 = "">
		<CFIF isDefined("FORM.Effectivity") AND Len(FORM.Effectivity) GT 100>
			<CFSET Effectivity2 = Right(FORM.Effectivity, Len(FORM.Effectivity) - 100)>
		</CFIF>


	<CFQUERY Name="UpdateFinding" DATASOURCE="#ODBC#">
		UPDATE TblAudit
		SET AuditDate = <CF_QUERYPARAM VALUE="#Form.AuditDate#" CFSQLTYPE="CF_SQL_DATE">,
			AuditType = <CF_QUERYPARAM VALUE="#Form.AuditType#" NULL="#isBlank(Form.AuditType)#" CFSQLTYPE="CF_SQL_VARCHAR">,
			FindingType = <CF_QUERYPARAM VALUE="#Form.FindingType#"  CFSQLTYPE="CF_SQL_VARCHAR">,
			Category = <CF_QUERYPARAM VALUE="#Form.Category#" NULL="#isBlank(Form.Category)#" CFSQLTYPE="CF_SQL_VARCHAR">,
			NumItems = <CF_QUERYPARAM VALUE="#VAL(Form.NumItems)#" CFSQLTYPE="CF_SQL_INTEGER">,
			<CFSET iRepeat = 0>
			<CFIF Form.RepeatItem is "yes">
				<CFSET iRepeat = 1>
			</CFIF>
			RepeatItem = <CF_QUERYPARAM VALUE="#VAL(iRepeat)#" CFSQLTYPE="CF_SQL_BIT">,
			Classification = <CF_QUERYPARAM VALUE="#Form.ClosureList#" NULL="#isBlank(Form.ClosureList)#" CFSQLTYPE="CF_SQL_VARCHAR">,
			COE = <CF_QUERYPARAM VALUE="#Form.COE#" NULL="#isBlank(Form.COE)#" CFSQLTYPE="CF_SQL_VARCHAR">,
			Bldg = <CF_QUERYPARAM VALUE="#Form.Bldg#" NULL="#isBlank(Form.Bldg)#" CFSQLTYPE="CF_SQL_VARCHAR">,
			<CFSET Form.Workstation = Left(Form.Workstation,50)>
			Workstation = <CF_QUERYPARAM VALUE="#Form.Workstation#" NULL="#isBlank(Form.Workstation)#" CFSQLTYPE="CF_SQL_VARCHAR">,
			<CFSET Form.AuditName = Left(Form.AuditName,100)>
			AuditName = <CF_QUERYPARAM VALUE="#Form.AuditName#" NULL="#isBlank(Form.AuditName)#" CFSQLTYPE="CF_SQL_VARCHAR">,
		<CFIF FORM.OrigResponPerson IS NOT Form.ResponPerson>
			ResponPerson = <CF_QUERYPARAM VALUE="#Form.ResponPerson#" NULL="#isBlank(Form.ResponPerson)#" CFSQLTYPE="CF_SQL_VARCHAR">,
		</CFIF>
		<CFIF variables.showCoResponPerson and (form.coResponPerson neq "") and FORM.OrigCoResponPerson IS NOT Form.CoResponPerson>
			CoResponPerson = <CF_QUERYPARAM VALUE="#Form.coResponPerson#" NULL="#isBlank(Form.coResponPerson)#" CFSQLTYPE="CF_SQL_VARCHAR">,
		</CFIF>
			ClosureDueDate = <CF_QUERYPARAM VALUE="#DateFormat(ClosureDueDate, "mm/dd/yyyy")#" CFSQLTYPE="CF_SQL_DATE">,
			Citation = <CF_QUERYPARAM VALUE="#Citation#" NULL="#isBlank(Replace(Form.Citation, "|", " ", "ALL"))#" CFSQLTYPE="CF_SQL_VARCHAR" MAXLENGTH="#variables.citationMaxLength#">,
			Description = <CF_QUERYPARAM VALUE="#Description#" NULL="#isBlank(Replace(Form.Description, "|", " ", "ALL"))#" CFSQLTYPE="CF_SQL_VARCHAR">,
			CorrectiveAction = <CF_QUERYPARAM VALUE="#CorrectiveAction#" NULL="#isBlank(Replace(Form.CorrectiveAction, "|", " ", "ALL"))#" CFSQLTYPE="CF_SQL_VARCHAR">,
			CloseComment = <CF_QUERYPARAM VALUE="#Form.CloseComment#" NULL="#isBlank(Form.CloseComment)#" CFSQLTYPE="CF_SQL_VARCHAR">,
			ContactPerson = <CF_QUERYPARAM VALUE="#Form.ContactPerson#" NULL="#isBlank(Form.ContactPerson)#" CFSQLTYPE="CF_SQL_VARCHAR">,
			ContactPhone = <CF_QUERYPARAM VALUE="#Form.ContactPhone#" NULL="#isBlank(Form.ContactPhone)#" CFSQLTYPE="CF_SQL_VARCHAR" MAXLENGTH="50">,
			Status = <CF_QUERYPARAM VALUE="#Form.Status#" NULL="#isBlank(Form.Status)#" CFSQLTYPE="CF_SQL_VARCHAR">,
			CloseDate = <CF_QUERYPARAM VALUE="#Form.CloseDate#" NULL="#isBlank(Form.CloseDate)#" CFSQLTYPE="CF_SQL_TIMESTAMP">,
			ClosePerson = <CF_QUERYPARAM VALUE="#Form.ClosePerson#" NULL="#isBlank(Form.ClosePerson)#" CFSQLTYPE="CF_SQL_VARCHAR">,
			RefType = <CF_QUERYPARAM VALUE="#Form.RefType#" NULL="#isBlank(Form.RefType)#" CFSQLTYPE="CF_SQL_VARCHAR" MAXLENGTH="100">,
			RefID = <CF_QUERYPARAM VALUE="#Form.RefID#" NULL="#isBlank(Form.RefID)#" CFSQLTYPE="CF_SQL_VARCHAR" MAXLENGTH="50">,
			UpdateDate = <CF_QUERYPARAM VALUE="#DateFormat(Now(), "mm/dd/yyyy")# #TimeFormat(Now(), "HH:mm:ss")#" CFSQLTYPE="CF_SQL_TIMESTAMP">,
			UpdateUser = <CF_QUERYPARAM VALUE="#sAccessName#" NULL="#isBlank(sAccessName)#" CFSQLTYPE="CF_SQL_VARCHAR">,
			UpdateHistory = <CF_QUERYPARAM VALUE="#History#" NULL="#isBlank(History)#" CFSQLTYPE="CF_SQL_LONGVARCHAR">

		<!--- 08/16/02 --->
		<!--- 10/27/03 --->
		<CFIF bVerifyByDisabled IS false>
			<CFIF VerifyBy IS "" OR VerifyBy IS NOT VerifyBy_Orig>
				, VerifyPerson = <CF_QUERYPARAM VALUE="#VerifyBy#" NULL="#isBlank(VerifyBy)#" CFSQLTYPE="CF_SQL_VARCHAR">
			</CFIF>
				, VerifyDate = <CF_QUERYPARAM VALUE="#VerifyDate#" NULL="#!isDate(VerifyDate)#" CFSQLTYPE="CF_SQL_DATE">
				, VerifyComment = <CF_QUERYPARAM VALUE="#VerifyComment#" NULL="#isBlank(VerifyComment)#" CFSQLTYPE="CF_SQL_VARCHAR">
		<cfelseif UpdateVerifyBy><!--- 6/3/04 --->
			, VerifyPerson = <CF_QUERYPARAM VALUE="#VerifyBy#" NULL="#isBlank(VerifyBy)#" CFSQLTYPE="CF_SQL_VARCHAR">
		</CFIF>
		<!--- 06/19/04 --->
		<CFIF FORM.Status IS "Open" AND bVerifyByDisabled IS true>
			, VerifyDate = <CF_QUERYPARAM VALUE="" NULL="Yes" CFSQLTYPE="CF_SQL_DATE">
		</CFIF>
		<!--- 08/29/02 --->
			,DaysBeforeReminder = <CF_QUERYPARAM VALUE="#DaysBeforeReminder#" NULL="#isBlank(DaysBeforeReminder)#" CFSQLTYPE="CF_SQL_VARCHAR">
		<!--- 04/03/03 --->
			, VerifyByDate = <CF_QUERYPARAM VALUE="#VerifyByDate#" NULL="#!isDate(VerifyByDate)#" CFSQLTYPE="CF_SQL_DATE">
			, SubCOEID = <CF_QUERYPARAM VALUE="#FORM.SubCOEID#" NULL="#isBlank(FORM.SubCOEID)#" CFSQLTYPE="CF_SQL_BIGINT">

			, MULTICC = <CF_QUERYPARAM VALUE="#FORM.MULT_CC#" NULL="#isBlank(FORM.MULT_CC)#" CFSQLTYPE="CF_SQL_VARCHAR">
		<!--- 06/17/04 --->
		<CFIF ATS_CAPA_ENABLED IS true>
				, InvestigationDetails = <CF_QUERYPARAM VALUE="#FORM.InvestigationDetails#" NULL="#isBlank(FORM.InvestigationDetails)#" CFSQLTYPE="CF_SQL_VARCHAR">
				, #RootCauseColumn# = <CF_QUERYPARAM VALUE="#Left(FORM.RootCause, 600)#" NULL="#isBlank(Left(FORM.RootCause, 600))#" CFSQLTYPE="CF_SQL_VARCHAR" MAXLENGTH="600">
				,#EffectiveDateColumn# = <CF_QUERYPARAM VALUE="#FORM.EffectiveDate#" NULL="#isBlank(FORM.EffectiveDate)#" CFSQLTYPE="CF_SQL_VARCHAR" BLANKNULL="true">
			<CFSET iCapa = 1>
			<CFIF FORM.CAPARequired NEQ 1>
				<CFSET iCapa = 0>
			</CFIF>
				, #CAPARequiredColumn# = <CF_QUERYPARAM VALUE="#VAL(iCapa)#" CFSQLTYPE="CF_SQL_BIT">

			, #EffectivityColumn1# = <CF_QUERYPARAM VALUE="#FORM.Effectivity#" NULL="#isBlank(Left(FORM.Effectivity, 100))#" CFSQLTYPE="CF_SQL_VARCHAR" MAXLENGTH="100" BLANKNULL="true">
				, #EffectivityColumn2# = <CF_QUERYPARAM VALUE="#Effectivity2#" NULL="#isBlank(Effectivity2)#" CFSQLTYPE="CF_SQL_VARCHAR" BLANKNULL="true">
				, #RootCause1Column# = <CF_QUERYPARAM VALUE="#FORM.RootCause1#" NULL="#isBlank(FORM.RootCause1)#" CFSQLTYPE="CF_SQL_VARCHAR">

				, #RootCause2Column# = <CF_QUERYPARAM VALUE="#FORM.RootCause2#" NULL="#isBlank(FORM.RootCause2)#" CFSQLTYPE="CF_SQL_VARCHAR">
		</CFIF>
		<CFIF bHasSubCat>
			,#SubCategoryColumn# = <CF_QUERYPARAM VALUE="#Form.subCat#" CFSQLTYPE="CF_SQL_VARCHAR">
		</CFIF>
		<CFIF RiskCategoryEnabled>
			, ClassificationType = <CF_QUERYPARAM VALUE="#VAL(Form.RiskCategory)#" CFSQLTYPE="CF_SQL_INTEGER">
		</CFIF>
		<!--- Auto Escalation fields --->
		<cfif structKeyExists(form, "AutoEscalation")> 
			, AutoEscalation = <CF_QUERYPARAM VALUE="#VAL(form.AutoEscalation)#" CFSQLTYPE="CF_SQL_BIT">
		</cfif>
			, ExternalSubmit = <CF_QUERYPARAM VALUE="#VAL(Form.ExternalSubmit)#" CFSQLTYPE="CF_SQL_BIT">

	<!---addtional fields --->
		<CFIF isDefined("additionalATSFields") AND isStruct(additionalATSFields)>
			<CFLOOP LIST="#StructKeyList(additionalATSFields)#" INDEX="additionalGroup">
				<CFLOOP ARRAY="#additionalATSFields[additionalGroup]#" INDEX="Field">
					<CFIF Field.Name IS NOT "" && (!structKeyExists(field, "extensionsData"))>
						,#Field.Name# = <CF_QUERYPARAM VALUE="#FORM[Field.Name]#" CFSQLTYPE="CF_SQL_VARCHAR">
					</CFIF>
				</CFLOOP>
			</CFLOOP>
		</CFIF>

		WHERE Orgname=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
			and Location = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.Location#">
			and  ID =<CF_QUERYPARAM VALUE="#Form.ID#" CFSQLTYPE="CF_SQL_BIGINT">
	</CFQUERY>

	<!--- if approvals feature is active lets, process the approvals for them.  --->

	<cfset variables['extensionRecordSet'] = {} />
	
	<cfloop collection="#form#" item = "element">
		<cfif findNoCase("approvals_", element)>
			<cfset variables['extensionRecordSet']["#element#"]	= form["#element#"] />
		</cfif>
	</cfloop>

	<cfif !structIsEmpty(variables['extensionRecordSet'])>
		<cfset variables['extensionRecordSet']["tblAuditId"] = form.tblAuditId />
		<cfset variables.ExtensionFieldList = StructKeyList(variables['extensionRecordSet'],",")/>
		<cfset variables.approvalExistingQuery = 
				variables.extension.getRecords(whereRecordName="tblauditId",whereRecordValue="#form.tblAuditId#",fieldlist="#variables.ExtensionFieldList#",dsn="#odbc#")/>
		<cfif variables.approvalExistingQuery.recordCount gt 0>
			<cfset variables.extension.updRecord(record_group_id = variables.approvalExistingQuery.record_group_id,scope_siteID=form.siteID,update_by = request.user.accessName,recordset = variables['extensionRecordSet'],insertIfNotExists = true)>
		<cfelse>
			<cfset variables.extension.insRecord(created_by=request.user.accessName,created_date = now(),scope_siteID=form.siteID,recordset =variables.extensionRecordSet)> 	
		</cfif>
	</cfif>
	<!--- ENDS approvals section on EDIT mode --->
			<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#additionalFields.cfm"
				  siteid="#form.siteid#"
				  FIELDS="additionalATSFields" 
				  reftype="ats"
				  odbc="#odbc#"
				  type="edit"
				  orgid    = "#form.orgid#"
				  refid="tblAuditID"
				  refidValue="#FORM.tblAuditID#"
				  datascope="#form#"
				  TRANSLATOR=""/>

	<cfif form.status eq "Open" and VerifyDate eq "">
		<!--- At this point finding was already verified. But was re-opened --->
		<cfmodule template="#request.library.customtags.virtualpath#extensionshandler.cfm"
				appid="#request.app.appID#" ref_id="#form.tblAuditID#" action="open"
				audithome="#audithome#" odbc="#odbc#" siteid="#locId#" />
	</cfif>

	<CFCATCH TYPE="Database">
		<CFIF Request.isProduction EQ 0>
			<CFRETHROW>
		</CFIF>
<!---
		<CFTRANSACTION ACTION="ROLLBACK"/>
--->

		<CFMAIL TO="ActionTrackingSystem.PM@gensuitellc.com" cc="" FROM="ActionTrackingSystem.PM@gensuitellc.com"
		 SUBJECT="audaction.cfm - Failure" TYPE="HTML">
		 	<CFDUMP VAR="#CGI#">
		 	<CFDUMP VAR="#CFCATCH#">
		</CFMAIL>

		<CFMODULE TEMPLATE="#Request.Library.framework.VirtualPath#tollgateMessage.cfm" TYPE="ERROR">
			<div id="error">Sorry! Either incorrect input or network and/or browser difficulties are preventing your inputs from being received completely for processing!</div><p>
			<cfif Request.IsBlackBerry EQ True>
			<!--- Added to Allow BlackBerry to go Back to Previous Page - CJ 10/11/2005 --->
			Please use the Go Back button to return to your input form and re-submit.
			<cfoutput>#thegobackform#</cfoutput>
			<!--- <cfelse>
			<cfoutput>#BACKLINK#</cfoutput>
			 to go back and re-submit your input form.  Please note that you may need to reload your form page in order to submit your inputs successfully. ---></font><BR><BR>
			</cfif>
			<cfif Request.IsBlackBerry NEQ True>
			<cfoutput>
			<A target="_top" HREF="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##ContactsHome#email.cfm">Contact</A> the business administrator if this problem persists.  Please be sure to use Microsoft Internet Explorer 5.0 or higher for best results.
			</cfoutput>
			</cfif>
			<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#PDA/PDASYNC.cfm" ACTION="Failure"></CFMODULE>
		</CFMODULE>
	</CFCATCH>
	</CFTRY>

</CFIF>

<CFIF FORM.Action IS NOT "Delete" AND FORM.Action IS NOT "View">


	<!--- Need to pass these to the redirect.. so add them --->
	<CFSET REDIRECTDATA.FORM.ID = Form.ID>
	<CFSET REDIRECTDATA.FORM.bAdded = Form.bAdded>
	<CFSET REDIRECTDATA.FORM.tblAuditID = Form.tblAuditID>
	<CFSET REDIRECTDATA.FORM.dateadded = now()>

	<CFSET sData = SerializeJSON(REDIRECTDATA)>
	<CFSET ActionID = Replace(CreateUUID(), "-", "_", "All")>
	
	
	
	<cftry>
		<CFQUERY NAME="qGetData" DATASOURCE="#ODBC#" RESULT="WTF">
			if not exists (select * from tempdb.sys.objects (nolock) where name='####ATS_REDIRECT_DATA' and type='U')
			    create table ####ATS_REDIRECT_DATA (
			        Data nvarchar(max),
					id varchar(35),
					updatedate datetime
			    );

			INSERT INTO ####ATS_REDIRECT_DATA
			(DATA, id, updatedate)
			VALUES
			(
			N'#sData#',
			<CF_QUERYPARAM VALUE="#ActionID#" CFSQLTYPE="CF_SQL_VARCHAR">,
			getdate()
			)
		</CFQUERY>
		<cfcatch type="any">
			<cfmail to='miguel.aguilar@gensuitellc.com' from='miguel.aguilar@gensuitellc.com' subject='ATS_REDIRECT_DATA - something failed.' type='HTML'>
				<cfdump var="#cfcatch#">
			</cfmail>
		</cfcatch>
	</cftry>
	<!---Neha comment begin for Child Action Item: 237534
	Neha comment end for child action item: 237534--->
	<!--- to pass removed RP email IDs through URL
	SHOULD BE FIXED... AND ACCESSED THROUGH THE FORM
	--->
	<cfif request.businessid neq 1305>
		<cfif structKeyExists(request,"CustomTaggerEnabled") and request.CustomTaggerEnabled eq true and variables.embedCustomTagger eq true>
	        <script type="text/javascript">
	            <cfoutput>
	                window.location.replace("#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#audaction.cfm?actionid=#actionid#&factoryid=#factoryid#");
	            </cfoutput>
	        </script>
	    <cfelse>
	        <CFLOCATION ADDTOKEN="No" URL="audaction.cfm?ActionID=#ActionID#&FactoryID=#FactoryID# "> 
	    </cfif>
	</cfif>
</CFIF>

<!--- 
<HTML>
<HEAD --->

	<!--- 
	<CFOUTPUT>
		<TITLE>#Abr# #REQUEST.ATSFindingName#</TITLE>
			
		</CFOUTPUT> --->
	
	
<script language="JavaScript">
	<CFINCLUDE TEMPLATE="#REQUEST.LIBRARY.JAVASCRIPT.VIRTUALPATH#gsWindowOpen.js">

	function popup(loc,winName) {
	    gsWindowOpen(loc, winName);
	}
</script>
<!--- </HEAD> --->
<!--- <BODY style="overflow:scroll !important"> --->
<DIV STYLE="margin:5px; width:99%;">

<CFSAVECONTENT VARIABLE="sMessage">
<CFSET variables.FindingID = "">
<CFIF #Form.Action# EQ "Add">
	<CFSET variables.FindingID = #NewID#>
<!--- 08/16/02 --->
<CFELSEIF Form.Action IS NOT "Delete">
	<CFSET variables.FindingID = #Form.ID#>
</CFIF>

<cfif not isDefined("addSimilar")>
<CFOUTPUT>
<CFIF FORM.OfflineMode IS "NO" AND FORM.PDAMODE EQ True>
<!--- Added Link Here for Pocket PC to back to ATS, since Header goes back to iP  - Chuck Jody 05/26/2004 --->
	<CFSET sPage = "audit.cfm">
	<CFSET PCView = "&PDAmode=true">
	<CFIF Form.PCView IS "yes" OR URL.PCView IS "Yes">
		<CFSET sImgWidth = "WIDTH=200">
		<CFSET sTarget = "target=_self">
		<CFSET sPage = "audfinding.cfm">
		<CFSET PCView = "&PDAmode=true&PCView=yes">
	</CFIF>
	    <TABLE WIDTH="100%" CELLPADDING="0" CELLSPACING="0" BORDER="0">
    <TR>
        <TD>
        <font class="navlink">
        <A HREF="#sCurrentURL##sPage#?Org=#Org#&Loc=#EncodeForURL(Loc)##PCView#&IsBlackBerry=#Request.IsBlackBerry#" onMouseOver="self.status=Click to return to the #request.ATSAppName# window';return true;" onMouseOut="self.status=''; return true" target="_self">Return to the #request.ATSAppName#...</A></font>
        </TD>
        </TR>
    </TABLE>
</CFIF>

<cfif FORM.PDAMODE NEQ True>
	<DIV CLASS="mediumtexti">
<CFELSE>
	<DIV CLASS="container">
		<div class="control-group">
		<ul style="text-align: right; list-style-type: none; font-style:normal; font-weight:normal; font-size:14px; padding:0; margin:0">
					<li><cfoutput><a href="#AuditHome#audfinding.cfm?Org=#Org#&ORgName=#EncodeForURL(OName)#&Loc=#EncodeForURL(Loc)#">Add New #REQUEST.ATSFindingName#</a></cfoutput></li>
					<li><cfoutput><a href="auditPDA.cfm?Org=#org#&loc=#location#">Search #REQUEST.ATSFindingName#s</a></cfoutput></li>
				</ul>
		</div>
</cfif>
<cfif !StructKeyExists(URL, "replicate")>
	<div class="alert alert-success" role="alert">
	<cfoutput>
		#translator.translate("#REQUEST.ATSFindingName# successfully")# <CFIF #FORM.firstAction# is "Add">#translator.translate("added")#<CFELSEIF #FORM.firstAction# is "Edit">#translator.translate("edited")#<CFELSEIF FORM.firstAction IS "Delete">#translator.translate("deleted")#<CFELSEIF FORM.firstAction IS "Reject">#translator.translate("rejected")#<CFELSE>#translator.translate("saved")#</CFIF>.
	</cfoutput>
	#sAdditionalMessage#
</div>
</cfif>
<cfoutput>
	
<cfif REQUEST.BrowserCheck.isWebkitMobile()>
	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#GetAttachments.cfm"
		query="qGetAttachments"
		Org="#Org#"
		Loc="#Loc#"
		AppName="#request.ATSAppName#"
		AttachParent="Audit"
		AttachID="#FORM.ID#"
		BaseAttachPath="#AuditHome#attachments"
		SiteID="#LocID#" >


	<cfmodule TEMPLATE="#Request.Library.CustomTags.VirtualPath#attachDisplay.cfm"
		qAttach="#qGetAttachments#"
		manageIcon="true"
		BusinessID="#BusinessID#"
		AppID="#Request.App.AppID#"
		SiteID="#LocID#"
		RefType="Audit"
		RefID="#FORM.ID#"
		ODBC="#CC_ODBC#"/>
	<BR>
</cfif>

<!-- SUCCESS ID [#Form.ID#] -->
<!-- SUCCESS LINK [#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#audfinding.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&ID=#FORM.ID#] -->
</cfoutput>


<!--- 08/16/02 --->
<CFQUERY NAME="RPEmail" DATASOURCE="#ODBC#">
	Select CONTACT_EMAIL
	FROM ltbContact WITH (NOLOCK)
	WHERE CONTACT_NAME = <CF_QUERYPARAM VALUE="#Form.ResponPerson#" CFSQLTYPE="CF_SQL_VARCHAR">
</CFQUERY>
<cfif variables.showCoResponPerson and (form.coResponPerson neq "")>
	<CFQUERY NAME="CoRPEmail" DATASOURCE="#ODBC#">
		Select CONTACT_EMAIL
		FROM ltbContact WITH (NOLOCK)
		WHERE CONTACT_NAME = <CF_QUERYPARAM VALUE="#Form.coResponPerson#" CFSQLTYPE="CF_SQL_VARCHAR">
	</CFQUERY>
</cfif>
<CFQUERY NAME="CPEmail" DATASOURCE="#ODBC#">
	Select CONTACT_EMAIL
	FROM ltbContact WITH (NOLOCK)
	WHERE CONTACT_NAME = <CF_QUERYPARAM VALUE="#Form.ContactPerson#" CFSQLTYPE="CF_SQL_VARCHAR">
</CFQUERY>

<CFSET StatusText = audit.GetStatusTextForDisplay(Status="#Form.Status#", CILINK="#CI_actions_comp#", VerifyDate="#VerifyDate#", VERIFYPERSON="#VerifyBy#", Action="#FORM.Action#", ReOpenStatus="#ReOpenEmail#", EmailStatusText="True")>

<!--- 11/06/02 --->
<CFIF FORM.firstAction IS "Reject" OR ReOpenEmail IS "yes">
	<CFQUERY NAME="qGetResponsibleEmail" DATASOURCE="#ODBC#">
		SELECT CONTACT_EMAIL
		FROM ltbContact WITH (NOLOCK)
		WHERE CONTACT_NAME = <CF_QUERYPARAM VALUE="#Form.ResponPerson#" CFSQLTYPE="CF_SQL_VARCHAR">
	</CFQUERY>
	<CFQUERY NAME="qGetContactEmail" DATASOURCE="#ODBC#">
		SELECT CONTACT_EMAIL
		FROM ltbContact WITH (NOLOCK)
		WHERE CONTACT_NAME = <CF_QUERYPARAM VALUE="#Form.ContactPerson#" CFSQLTYPE="CF_SQL_VARCHAR">
	</CFQUERY>
	<!--- 10/01/2007 --->
	<CFIF Form.VerifyBy IS "">
		<CFSET Form.VerifyBy = AccessName>
	</CFIF>
	<CFQUERY NAME="qGetVerifyEmail" DATASOURCE="#ODBC#">
		SELECT CONTACT_EMAIL
		FROM ltbContact WITH (NOLOCK)
		WHERE CONTACT_NAME = <CF_QUERYPARAM VALUE="#Form.VerifyBy#" CFSQLTYPE="CF_SQL_VARCHAR">
	</CFQUERY>

	<!--- mail the Responsible Person to let them know that it was reopend. --->
	<!--- 11/06/02 --->
	<CFSET CCMail = qGetContactEmail.CONTACT_EMAIL>
	<CFIF FORM.firstAction IS NOT "Reject">
		<CFSET FromMail = EMailSender>
	<CFELSE>
		<CFSET FromMail = qGetVerifyEmail.CONTACT_EMAIL>
		<CFSET CCMail = ListAppend(CCMail, qGetVerifyEmail.CONTACT_EMAIL, ",")>
	</CFIF>

	<!--- 04/20/04 --->
	<CFSET CCList = ListAppend(CCMail, FORM.Mult_CC)>
	<cfif variables.showCoResponPerson and (form.coResponPerson neq "")>
		<CFSET CCList = ListAppend(CCList, CoRPEmail.CONTACT_EMAIL)>
	</cfif>
	<cfif qGetResponsibleEmail.CONTACT_EMAIL NEQ "">
	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#CLEANEMAIL.cfm"
		ToList = "#qGetResponsibleEmail.CONTACT_EMAIL#"
		CCList = "#CCList#"
		Delim = ",">
	</cfif>
	<cfif NOT IsDefined("ToList")><cfset ToList="#PowerSuiteAdminEmail#"></cfif>
	<cfif ToList EQ "">           <cfset ToList="#PowerSuiteAdminEmail#"></cfif>
	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#CLEANEMAIL.cfm"
		ToList = "#ToList#"
		CCList = "#CCList#"
		Delim = ",">

	<!--- "[FindingType] Type Finding Notification" distribution lists setup in contacts permissions if activated for this business   --->
	<cfmodule template="#ContactsMasterPath#distGetList.cfm"
		Application = "#request.ATSAppName#"
	    List = "#Form.FindingType# Type Finding Notification"
		AdditionalToEmails="#ToList#"
		AdditionalCcEmails="#CCList#"
	    SiteID ="#LocID#"
		DSN="#ODBC#">
	<cfset ToList = DistList.ToEmails>
	<cfset CCList = DistList.ccEmails>

	<CFSAVECONTENT VARIABLE="EmailNotes">
		<CFOUTPUT>
		<CFIF FORM.firstAction IS NOT "Reject">
			<CFSET VerificationAction = "Re-Opened">
			<b>Dear <font color="blue">
				<i>#Form.ResponPerson#:</b></i>
			</font><br>
			The #LCase(REQUEST.ATSAuditName)# #LCase(REQUEST.ATSFindingName)# below has been reopened.  You are
			 designated as the person responsible for leading its closure.<BR>

			Summary #LCase(REQUEST.ATSFindingName)# details are provided below.
		<CFELSE>
			<cfif CLosureVerificationlabelForEmails contains "verification">
				<CFSET VerificationAction = "Closure Rejected">
			<cfelse>
				<CFSET VerificationAction = "#CLosureVerificationlabelForEmails# Rejected">
			</cfif>
			<b>Dear <font color="blue">
				<i>#Form.ResponPerson#:</b></i>
			</font><br>
			Closure Verification for the #LCase(REQUEST.ATSAuditName)# #LCase(REQUEST.ATSFindingName)# below has been completed by #AccessName#
			and
			<FONT COLOR="red">
				closure has been <STRONG>rejected</STRONG>
			</FONT>
			because "<i>#EncodeForHTML(Form.VerifyComment)#</i>".
			<p>
			This #LCase(REQUEST.ATSFindingName)# has been <i>re-opened</i> with the original #LCase(REQUEST.ATSFindingName)# date...please see details below.  Please contact the designated Closure Verification person for follow-up.
		</CFIF>
		</CFOUTPUT>
	</CFSAVECONTENT>

	<cfif findingTypeLabel eq "">
		<cfset SubjectLine = "#Abr# #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# ##" & FORM.ID & " #LCase(VerificationAction)#: " & OName & ", " & EncodeForHTML(FactoryID) & Loc & " (#GEBusiness#)">
	<cfelse>
		<CFSET SubjectLine = "#Form.FindingType# #Abr# #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# ##" & FORM.ID & " #LCase(VerificationAction)#: " & OName & ", " & EncodeForHTML(FactoryID) & Loc & " (#GEBusiness#)">
	</cfif>
	<!--- Finding Re-opened
	(missing from Online Help)
	--->
	<CFSET inFile = "">
	<cfset FileName = "ATS_Finding_#ID#_" & RandRange(0,9999)>

	<!--- Write File: Task Closure form to send back to Task Owner after All SubTasks are Closed- CJ 05/03/2007 --->
	<CFIF FORM.ExternalSubmit EQ 1>
		<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#WriteOfflinePage.cfm"
			Action="WriteFile"
			PageURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#status.cfm"
			FileName="#FileName#"
			SiteID="#LocID#"
			ID="#ID#"
			ExternalSubmitForm="true"
			AttachedFileName="inFile">

		<CFIF inFile IS NOT "">
			<CFSAVECONTENT VARIABLE="Footer">
				<CFOUTPUT>#Footer#</CFOUTPUT>
				<!--- External Submit --->
				<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#ExternalSubmit.cfm"
						Action="ExternalSubmitInfo"/>
			</CFSAVECONTENT>
		</CFIF>
	</CFIF>

	
		<CFMODULE TEMPLATE="#REQUEST.LIBRARY.CUSTOMTAGS.VIRTUALPATH#EmailDetails.cfm"
		 	Details="#Details#"
			DetailsTitle="#REQUEST.ATSFindingName# Details"
		 	Summary="#DetailsSummary#"
			SummaryTitle="#REQUEST.ATSFindingName# Summary"
			ApplicationName="#request.ATSAppName#"
			ApplicationIcon="ats.gif"
			ApplicationURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#audit.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&FactoryID=#EncodeForURL(factoryid)#"
			OrganizationName="#OName#"
			SiteName="#factoryid##Form.Location#"
			BlockTitle="#REQUEST.ATSFindingName# ID##"
			BlockID="#Form.ID#"
			BlockIDURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#status.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&ID=#EncodeForURL(FORM.ID)#&FactoryID=#EncodeForURL(factoryid)#"
			EmailType="#REQUEST.ATSAuditName# #REQUEST.ATSFindingName# Summary"
			EmailNotes="#EmailNotes#"
			StatusTitle="Status" StatusText="#StatusText#"
			LINKURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#audfinding.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&ID=#EncodeForURL(FORM.ID)#&FactoryID=#EncodeForURL(factoryid)#"
			LINKTEXT="Edit this #REQUEST.ATSAuditName# #REQUEST.ATSFindingName#"
			TO="#ToList#"
			CC="#CCList#"
			<!--- to="diego.serrano@gensuitellc.com" --->
			FROM="#FromMail#"
			SUBJECT="#SubjectLine#"
			MailParam="X-Priority:1,X-MSMail-Priority:High,X-Message-Flag:Follow Up, Reply-By:#DateFormat(ClosureDueDate, 'ddd, d, mmm yyyy')# 12:00:00 -0500"
			FOOTER="#Footer#"
			MimeAttach="#inFile#"
			>
  
		<p>
		<font class="extramediumtext"><i>
		<CFIF ToList EQ "">
			The designated Responsible Person does not have a valid email address, therefore no email was sent.
		<CFELSE>
			<cfoutput>
			This email is being sent to the designated Responsible Person at <b>#qGetResponsibleEmail.CONTACT_EMAIL#</b><CFIF qGetContactEmail.CONTACT_EMAIL IS NOT ""> with cc to the Auditor/Contact Person at <b>#qGetContactEmail.CONTACT_EMAIL#</b></CFIF>!
			</cfoutput>
		</CFIF>
		</i></font>
<!--- 05/28/03 --->
<CFELSEIF IsDefined("Form.EmailRP") AND Form.Status IS NOT "Closed" AND FORM.firstAction IS NOT "Delete" AND NOT IsDefined("EmailAlreadySent")>

	<!--- Send email check box enabled AND Status is open--->
	<CFIF RPEmail.CONTACT_EMAIL NEQ "">

	<!--- 04/20/04 --->
	<CFSET CCList = ListAppend(CCList, ListAppend(CPEmail.CONTACT_EMAIL, FORM.Mult_CC))>
	<CFSET ToList = ListAppend(Tolist, RPEmail.CONTACT_EMAIL)>
	<cfif variables.showCoResponPerson and (form.coResponPerson neq "") and isDefined("CoRPEmail.CONTACT_EMAIL")>
		<CFSET ToList = ListAppend(Tolist, CoRPEmail.CONTACT_EMAIL)>
	</cfif>
	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#CLEANEMAIL.cfm"
		ToList = "#ToList#"
		CCList = "#CCList#"
		Delim = ",">

	<!---Neha change begin for child action item: 237534--->
	<!---To check for new added RP email Ids--->
	<CFIF EnableSubCorrectiveActions IS true and isDefined("FORM.SubCALastCount")>
		<CFLOOP FROM="#FORM.SubCALastCount#" TO="#FORM.SubCAAddCount-1#" INDEX="iSubCACount">
			<CFPARAM NAME="FORM.ResponPerson#iSubCACount#" DEFAULT="">
			<CFPARAM NAME="FORM.Description#iSubCACount#" DEFAULT="">
			<CFIF FORM["ResponPerson#iSubCACount#"] IS NOT "">
				<CFQUERY NAME="qGetSubEmail" DATASOURCE="#ODBC#">
					SELECT Contact_Name, Contact_Email
					FROM ltbContact WITH (NOLOCK)
					WHERE Contact_Name = <CF_QUERYPARAM VALUE="#FORM["ResponPerson#iSubCACount#"]#" CFSQLTYPE="CF_SQL_VARCHAR"/>
				</CFQUERY>
				<CFSET CCcount2 = 0>
				<CFIF FORM["RESPONPERSON#iSubCACount#"] IS NOT "">
					 <CFLOOP LIST="#CCList#" INDEX="icheck">
						 <CFIF ("CCList#icheck#" NEQ qGetSubEmail.Contact_Email)>
						 	<CFSET CCcount2 = CCcount2 +1>
						 <CFELSE>
						 	<CFSET CCcount2 = -1>
							<CFBREAK>
						</CFIF>
					</CFLOOP>
				 </CFIF>
				<CFIF CCcount2 NEQ -1>
					<CFSET CCList = ListAppend(CCList, qGetSubEmail.Contact_Email)>
				</CFIF>
			</CFIF>
		</CFLOOP>
	</cfif>
	<CFIF FORM.CCDeleteList IS NOT ''>
		<CFSET CCList = ListAppend(CCList, FORM.CCDeleteList)>
	</CFIF>
	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#CLEANEMAIL.cfm"
	ToList = "#ToList#"
	CCList = "#CCList#"
	Delim = ",">

	<!--- "[FindingType] Type Finding Notification" distribution lists setup in contacts permissions if activated for this business   --->
	<cfmodule template="#ContactsMasterPath#distGetList.cfm"
		Application = "#request.ATSAppName#"
	    List = "#Form.FindingType# Type Finding Notification"
		AdditionalToEmails="#ToList#"
		AdditionalCcEmails="#CCList#"
	    SiteID ="#LocID#"
		DSN="#ODBC#">
	<cfset ToList = DistList.ToEmails>
	<cfset CCList = DistList.ccEmails>

	<cfset ltvVErifierEmail = ""/>
	<cfif trim(ltv_verifier) neq "">
		<cfquery name="qltvEmail" datasource="#odbc#">
			select top 1 Contact_Email
			FROM ltbContact WITH (NOLOCK)
						WHERE Contact_Name = <CF_QUERYPARAM VALUE="#ltv_verifier#" CFSQLTYPE="CF_SQL_VARCHAR"/>
		</cfquery>
		<cfset ltvVErifierEmail = qltvEmail.Contact_Email />
		<cfset CCList = listAppend(cclist, ltvVErifierEmail, ",")/>
	</cfif>

	<!---Neha change end for Child Action Item: 237534--->
	<CFSAVECONTENT VARIABLE="EmailNotes">
		<CFOUTPUT>
		<b>Dear <font color="blue"><i>#Form.ResponPerson#:</b></i></font><br>
				<CFIF FORM.bAdded>
				<CFSET EmailAction = "added">
				The #LCase(REQUEST.ATSAuditName)# #LCase(REQUEST.ATSFindingName)# below has been added and you are designated as the person responsible for leading its closure.
				<cfif structKeyExists(form, "selectedSiteIdList") && len(trim(form.selectedSiteIdList))>
					<br><br><b>Note:</b>&nbsp;Email generated immediately for site where initial finding was logged.  For all replicated sites, emails will be sent during next upcoming batch (runs thrice a day at 12:00 AM ET, 08:00 AM ET and 04:00 PM ET).
				</cfif>
				<cfelse>
				<CFSET EmailAction = "updated">
				The #LCase(REQUEST.ATSAuditName)# #LCase(REQUEST.ATSFindingName)# below has been updated and you are designated as the person responsible for leading its closure.
				</CFIF>
		</CFOUTPUT>
	</CFSAVECONTENT>

	 <cfif findingTypeLabel eq "">
	 	<cfset SubjectLine = "#Abr# #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# ##" & #variables.FindingID# & " #EmailAction#: " & OName & ", " & EncodeForHTML(FactoryID) & Loc & " (#GEBusiness#)">
	 <cfelse>
	 	<CFSET SubjectLine = "#Form.FindingType# #Abr# #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# ##" & #variables.FindingID# & " #EmailAction#: " & OName & ", " & EncodeForHTML(FactoryID) & Loc & " (#GEBusiness#)">
	 </cfif>
	 <CFIF Request.user.accessemail is not "">
	 	<cfset FromEmail = "#Request.user.accessemail#">
	 <CFELSE>
	 	<cfset FromEmail = "#EMailSender#">
	 </CFIF>


	<!--- this email does't show the RP since it is to the RP --->
	<CFSET ArrayDeleteAt(Details, iResponsiblePersonLocation)>
	<!--- Add/Edit Email
	http://cincep09corpge.corporate.ge.com/help/ccenter/ats/ATS_Finding.mht
	--->
	<CFSET inFile = "">
	<cfset FileName = "ATS_Finding_#ID#_" & RandRange(0,9999)>

	<!--- Write File: Task Closure form to send back to Task Owner after All SubTasks are Closed- CJ 05/03/2007 --->
		<CFIF FORM.ExternalSubmit EQ 1>
			<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#WriteOfflinePage.cfm"
				Action="WriteFile"
				PageURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#status.cfm"
				FileName="#FileName#"
				SiteID="#LocID#"
				ID="#ID#"
				ExternalSubmitForm="true"
				AttachedFileName="inFile">

			<CFIF inFile IS NOT "">
				<CFSAVECONTENT VARIABLE="Footer">
					<CFOUTPUT>#Footer#</CFOUTPUT>
					<!--- External Submit --->
					<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#ExternalSubmit.cfm"
							Action="ExternalSubmitInfo"/>
				</CFSAVECONTENT>
			</CFIF>
		</CFIF>

		<cfquery name="qCheckEmail" datasource="#odbc#">
			select 	id
			from 	incident_data with (nolock)
			where 	reftype='ATSSummaryEmails'
			    and refid='#variables.findingID#'
			    and isdate(isnull(fieldvalue,'')) = 1
				and DateDiff(d,fieldvalue, GetDate())=0
		</cfquery>

		<cfif not qCheckEmail.recordcount && form.replication eq false>
			<cfquery name="insertEmailSentRecord" datasource="#odbc#">
				insert into incident_data
						(reftype,
						refid, 
						fieldname, 
						fieldvalue)
				values ('ATSSummaryEmails', '#variables.findingID#', 'atsSummaryEmailSent',#createODBCDate(now())#)
			</cfquery>
		 	<!--- Open Finding --->
			<CFMODULE TEMPLATE="#REQUEST.LIBRARY.CUSTOMTAGS.VIRTUALPATH#EmailDetails.cfm"
			 	Details="#Details#"
				DetailsTitle="#REQUEST.ATSFindingName# Details"
			 	Summary="#DetailsSummary#"
				SummaryTitle="#REQUEST.ATSFindingName# Summary"
				ApplicationName="#request.ATSAppName#"
				ApplicationIcon="ats.gif"
				ApplicationURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#audit.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&FactoryID=#EncodeForURL(factoryid)#"
				OrganizationName="#OName#"
				SiteName="#factoryid##Form.Location#"
				BlockTitle="#REQUEST.ATSFindingName# ID##"
				BlockID="#Form.ID#"		BlockIDURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#status.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&ID=#EncodeForURL(FORM.ID)#&FactoryID=#EncodeForURL(factoryid)#"
				EmailType="#REQUEST.ATSAuditName# #REQUEST.ATSFindingName# Summary"
				EmailNotes="#EmailNotes#"
				StatusTitle="Status" StatusText="#StatusText#"
				LINKURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#audfinding.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&ID=#EncodeForURL(FORM.ID)#&FactoryID=#EncodeForURL(factoryid)#"
				LINKTEXT="Edit this #REQUEST.ATSAuditName# #REQUEST.ATSFindingName#"
				TO="#ToList#"
				CC="#CCList#"
				<!--- to="diego.serrano@gensuitellc.com" --->
				FROM="#FromEmail#"
				SUBJECT="#SubjectLine#"
				MailParam="X-Priority:1,X-MSMail-Priority:High,X-Message-Flag:Follow Up, Reply-By:#DateFormat(ClosureDueDate, 'ddd, d, mmm yyyy')# 12:00:00 -0500"
				FOOTER="#Footer#"
				MimeAttach="#inFile#"
				>
				<p>
					<font class="extramediumtext">
						<i>
				<CFIF ToList EQ "">
					The designated Responsible Person does not have a valid email address, therefore no email was sent.
				<CFELSE>
					<cfoutput>
					#translator.translate("An email on this #LCase(REQUEST.ATSAuditName)# #LCase(REQUEST.ATSFindingName)# has been sent as requested to the designated Responsible Person<CFIF ListLen(ToList) GT 1>(s)</CFIF> at")#  <b>#ToList#</b><CFIF CCList NEQ ""> #translator.translate("with cc to")# <b>#CCList#</b></CFIF>!
					</cfoutput>
								<cfif structKeyExists(form, "selectedSiteIdList") && len(trim(form.selectedSiteIdList))>
									<br><br>Note:&nbsp;Email generated immediately for site where initial finding was logged.  For all replicated sites, emails will be sent during next upcoming batch (runs thrice a day at 12:00 AM ET, 08:00 AM ET and 04:00 PM ET). 
								</cfif>
				</CFIF>
						</i>
					</font>
			</cfif> <!--- CHECK IF EMAIL NEEDS TO BE SENT TODAY, TEEHEE --->
		<CFELSE>
			<p>
			<font class="extramediumtext"><i>
			<cfoutput>
			Sorry, an email on this #LCase(REQUEST.ATSAuditName)# #LCase(REQUEST.ATSFindingName)# could not be sent as requested to the designated Responsible Person because an email address was not available in the #Abr# Contacts Database for <u>#Form.ResponPerson#</u>!
			</cfoutput>
			</i></font>
		</CFIF>
</CFIF>

<!--- send out closure email only if not being verified and the email checkbox is been checked --->
<CFIF FORM.Status IS "Closed" AND FORM.firstAction IS NOT "Delete" AND VerifyBy IS "" AND IsDefined("Form.EmailRP") AND Form.EmailRP IS NOT "">

	<!--- get the RP, CC, MultiCC, and Person who closed, and verifyby --->
	<CFSET ToList = "">
	<CFSET CCList = "">

	<CFQUERY NAME="qGetResponsibleEmail" DATASOURCE="#ODBC#">
		SELECT CONTACT_EMAIL
		FROM ltbContact WITH (NOLOCK)
		WHERE Contact_Name = <CF_QUERYPARAM VALUE="#Form.ResponPerson#" CFSQLTYPE="CF_SQL_VARCHAR">
	</CFQUERY>
	<CFIF qGetResponsibleEmail.CONTACT_EMAIL IS "">
		<CFQUERY NAME="qGetResponsibleEmail" DATASOURCE="#ODBC#">
			SELECT CONTACT_EMAIL
			FROM ltbContact WITH (NOLOCK)
			WHERE Contact_Name = <CF_QUERYPARAM VALUE="#AccessName#" CFSQLTYPE="CF_SQL_VARCHAR">
		</CFQUERY>
	</CFIF>
	<CFSET ToList = ListAppend(ToList, qGetResponsibleEmail.CONTACT_EMAIL)>
	<cfif variables.showCoResponPerson and (form.coResponPerson neq "")>
		<CFQUERY NAME="qGetcoResponsibleEmail" DATASOURCE="#ODBC#">
			SELECT CONTACT_EMAIL
			FROM ltbContact WITH (NOLOCK)
			WHERE Contact_Name = <CF_QUERYPARAM VALUE="#Form.coResponPerson#" CFSQLTYPE="CF_SQL_VARCHAR">
		</CFQUERY>
		<CFSET ToList = ListAppend(ToList, qGetcoResponsibleEmail.CONTACT_EMAIL)>
	</cfif>
	<CFQUERY NAME="qGetResponsibleEmailCC" DATASOURCE="#ODBC#">
		SELECT CONTACT_EMAIL
		FROM ltbContact WITH (NOLOCK)
		WHERE Contact_Name = <CF_QUERYPARAM VALUE="#Form.ContactPerson#" CFSQLTYPE="CF_SQL_VARCHAR">
	</CFQUERY>
	<CFSET CCList = ListAppend(CCList, qGetResponsibleEmailCC.CONTACT_EMAIL)>
	<CFIF FORM.MULT_CC IS NOT "">
		<CFSET CCList = ListAppend(CCList,FORM.MULT_CC)>
	</CFIF>
	<CFIF FORM.ClosePerson IS NOT AccessName>
		<CFQUERY NAME="qGetResponsibleEmail" DATASOURCE="#ODBC#">
			SELECT CONTACT_EMAIL
			FROM ltbContact WITH (NOLOCK)
			WHERE Contact_Name = <CF_QUERYPARAM VALUE="#AccessName#" CFSQLTYPE="CF_SQL_VARCHAR">
		</CFQUERY>
		<cfset fromEmail = "#qGetResponsibleEmail.CONTACT_EMAIL#">
	<CFELSE>
		<cfif request.user.AccessEmail is not "">
			<cfset fromEmail = "#request.user.AccessEmail#">
		<cfelse>
			<cfset fromEmail = "#EmailSender#">
		</cfif>

	</CFIF>

	<cfif findingTypeLabel eq "">
		<cfset SubjectLine = "#Abr# #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# ##" & FORM.ID & " has been closed: " & OName & ", " & EncodeForHTML(FactoryID) & Loc & " (#GEBusiness#)">
	<cfelse>
		<CFSET SubjectLine = "#EncodeForHTML(Form.FindingType)# #Abr# #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# ##" & FORM.ID & " has been closed: " & OName & ", " & EncodeForHTML(FactoryID) & Loc & " (#GEBusiness#)">
	</cfif>

	  <CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#CLEANEMAIL.cfm"
	  ToList = "#ToList#"
	  CCList = "#CCList#"
	  Delim = ",">

	<!--- "[FindingType] Type Finding Notification" distribution lists setup in contacts permissions if activated for this business   --->
	<cfmodule template="#ContactsMasterPath#distGetList.cfm"
		Application = "#request.ATSAppName#"
	    List = "#Form.FindingType# Type Finding Notification"
		AdditionalToEmails="#ToList#"
		AdditionalCcEmails="#CCList#"
	    SiteID ="#LocID#"
		DSN="#ODBC#">
	<cfset ToList = DistList.ToEmails>
	<cfset CCList = DistList.ccEmails>

	<!--- Finding Closed Email
	http://cincep09corpge.corporate.ge.com/help/ccenter/ats/closure.mht
	--->
	<CFMODULE TEMPLATE="#REQUEST.LIBRARY.CUSTOMTAGS.VIRTUALPATH#EmailDetails.cfm"
	 	Details="#Details#"
		DetailsTitle="#REQUEST.ATSFindingName# Details"
	 	Summary="#DetailsSummary#"
		SummaryTitle="#REQUEST.ATSFindingName# Summary"
		ApplicationName="#request.ATSAppName#"
		ApplicationIcon="ats.gif"
		ApplicationURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#audit.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&FactoryID=#EncodeForURL(factoryid)#"
		OrganizationName="#OName#"
		SiteName="#factoryid##Form.Location#"
		BlockTitle="#REQUEST.ATSFindingName# ID##"
		BlockID="#Form.ID#"
		BlockIDURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#status.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&ID=#EncodeForURL(FORM.ID)#&FactoryID=#EncodeForURL(factoryid)#"
		EmailType="#REQUEST.ATSAuditName# #REQUEST.ATSFindingName# Summary"
		EmailNotes="The following #LCase(REQUEST.ATSFindingName)# has been closed by #AccessName#.<BR>This email has been sent for informational purposes only, no further action is required at this time."
		StatusTitle="Status" StatusText="#StatusText#" StatusTextColor="blue"
		LINKURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#audfinding.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&ID=#EncodeForURL(FORM.ID)#&FactoryID=#EncodeForURL(factoryid)#"
		LINKTEXT="Edit this #REQUEST.ATSAuditName# #REQUEST.ATSFindingName#"
		 TO="#ToList#" 
		 CC="#CCList#" 
		<!--- to="diego.serrano@gensuitellc.com" --->
		FROM="#fromEmail#"
		SUBJECT="#SubjectLine#"
		FOOTER="#Footer#"
		>

		<p>
		<P>
		<FONT CLASS="extramediumtext">
		<I>
		<CFIF ToList EQ "">
			The designated Responsible Person does not have a valid email address, therefore no email was sent.
		<CFELSE>
			<CFOUTPUT>
			An email on this #LCase(REQUEST.ATSAuditName)# #LCase(REQUEST.ATSFindingName)# has been sent to the designated person to verify this
			 #LCase(REQUEST.ATSAuditName)# at <b>#ToList#</b>
			<CFIF RPEmail.CONTACT_EMAIL IS NOT "">
				with cc to the Auditor/Contact Person at <b>#CCList#</b>
			</CFIF>!
			</CFOUTPUT>
		</CFIF>
		</I>
		</FONT>

</CFIF>

<!--- send an email to indicate Closure Verification has been completed --->
<CFPARAM NAME="FORM.VerifyDateOrig" DEFAULT="#FORM.VerifyDate#">
<CFIF Form.Status IS "Closed" AND FORM.firstAction IS NOT "Delete" AND FORM.VerifyBy IS NOT "" AND FORM.VerifyDate IS NOT "" AND FORM.VerifyComment IS NOT "" AND FORM.VerifyDate IS NOT FORM.VerifyDateOrig AND isDefined("Form.EmailRP") AND Form.EmailRP IS NOT "">

	<CFSET ToList = "">
	<CFSET CCList = "">

	<CFQUERY NAME="qGetResponsibleEmail" DATASOURCE="#ODBC#">
		SELECT CONTACT_EMAIL
		FROM ltbContact WITH (NOLOCK)
		WHERE Contact_Name = <CF_QUERYPARAM VALUE="#Form.ResponPerson#" CFSQLTYPE="CF_SQL_VARCHAR">
	</CFQUERY>
	<CFIF qGetResponsibleEmail.CONTACT_EMAIL IS "">
		<CFQUERY NAME="qGetResponsibleEmail" DATASOURCE="#ODBC#">
			SELECT CONTACT_EMAIL
			FROM ltbContact WITH (NOLOCK)
			WHERE Contact_Name = <CF_QUERYPARAM VALUE="#AccessName#" CFSQLTYPE="CF_SQL_VARCHAR">
		</CFQUERY>
	</CFIF>
	<CFSET ToList = ListAppend(ToList, qGetResponsibleEmail.CONTACT_EMAIL)>
	<cfif variables.showCoResponPerson and (form.coResponPerson neq "")>
		<CFQUERY NAME="qGetcoResponsibleEmail" DATASOURCE="#ODBC#">
			SELECT CONTACT_EMAIL
			FROM ltbContact WITH (NOLOCK)
			WHERE Contact_Name = <CF_QUERYPARAM VALUE="#Form.coResponPerson#" CFSQLTYPE="CF_SQL_VARCHAR">
		</CFQUERY>
		<CFSET ToList = ListAppend(ToList, qGetcoResponsibleEmail.CONTACT_EMAIL)>
	</cfif>
	<CFQUERY NAME="qGetResponsibleEmailCC" DATASOURCE="#ODBC#">
		SELECT CONTACT_EMAIL
		FROM ltbContact WITH (NOLOCK)
		WHERE Contact_Name = <CF_QUERYPARAM VALUE="#Form.ContactPerson#" CFSQLTYPE="CF_SQL_VARCHAR">
	</CFQUERY>
	<CFSET CCList = ListAppend(CCList, qGetResponsibleEmailCC.CONTACT_EMAIL)>
	<CFIF FORM.MULT_CC IS NOT "">
		<CFSET CCList = ListAppend(CCList, FORM.MULT_CC)>
	</CFIF>
	<CFQUERY NAME="qGetResponsibleEmail" DATASOURCE="#ODBC#">
		SELECT CONTACT_EMAIL
		FROM ltbContact WITH (NOLOCK)
		WHERE Contact_Name = <CF_QUERYPARAM VALUE="#FORM.VerifyBy#" CFSQLTYPE="CF_SQL_VARCHAR">
	</CFQUERY>
	<CFSET CCList = ListAppend(CCList, qGetResponsibleEmail.CONTACT_EMAIL)>
	<CFSET FromEmail = qGetResponsibleEmail.CONTACT_EMAIL>

	<CFIF ToList IS NOT "">
	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#CLEANEMAIL.cfm"
			ToList = "#ToList#"
			CCList = "#CCList#"
			Delim = ",">

		<CFSAVECONTENT VARIABLE="EmailNotes">
			<CFOUTPUT>
				<b>Dear <font color="blue"><i>#FORM.ResponPerson#:</b></i></font>
				<br>
				The closure of the #LCase(REQUEST.ATSAuditName)# #LCase(REQUEST.ATSFindingName)# below has been accepted and verified. No further action is required at this time.
			</CFOUTPUT>
		</CFSAVECONTENT>

		<cfif findingTypeLabel eq "">
			<cfset SubjectLine = "#Abr#  #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# ##" & FORM.ID & " #CLosureVerificationlabelForEmails# complete: " & OName & ", " & EncodeForHTML(FactoryID) & Loc & " (#GEBusiness#)">
		<cfelse>
			<CFSET SubjectLine = "#EncodeForHTML(Form.FindingType)# #Abr#  #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# ##" & FORM.ID & " #CLosureVerificationlabelForEmails# complete: " & OName & ", " & EncodeForHTML(FactoryID) & Loc & " (#GEBusiness#)">
		</cfif>

		<!--- "[FindingType] Type Finding Notification" distribution lists setup in contacts permissions if activated for this business   --->
		<cfmodule template="#ContactsMasterPath#distGetList.cfm"
			Application = "#request.ATSAppName#"
		    List = "#Form.FindingType# Type Finding Notification"
			AdditionalToEmails="#ToList#"
			AdditionalCcEmails="#CCList#"
		    SiteID ="#LocID#"
			DSN="#ODBC#">
		<cfset ToList = DistList.ToEmails>
		<cfset CCList = DistList.ccEmails>

		<!--- Send Closure Completion email
		http://cincep09corpge.corporate.ge.com/help/ccenter/ats/Closure_Verification_Completed.mht
		--->
		<CFMODULE TEMPLATE="#REQUEST.LIBRARY.CUSTOMTAGS.VIRTUALPATH#EmailDetails.cfm"
		 	Details="#Details#"
			DetailsTitle="#REQUEST.ATSFindingName# Details"
		 	Summary="#DetailsSummary#"
			SummaryTitle="#REQUEST.ATSFindingName# Summary"
			ApplicationName="#request.ATSAppName#"
			ApplicationIcon="ats.gif"
			ApplicationURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#audit.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&FactoryID=#EncodeForURL(factoryid)#"
			OrganizationName="#OName#"
			SiteName="#factoryid##Loc#"
			BlockTitle="#REQUEST.ATSFindingName# ID##"
			BlockID="#ID#"
			BlockIDURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#status.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&ID=#ID#&FactoryID=#EncodeForURL(factoryid)#"
			EmailType="#REQUEST.ATSAuditName# #REQUEST.ATSFindingName# Summary"
			EmailNotes="#EmailNotes#"
			StatusTitle="Status" StatusText="#StatusText#" StatusTextColor="blue"
			LINKURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#audfinding.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&ID=#ID#&FactoryID=#EncodeForURL(factoryid)#"
			LINKTEXT="Edit this #REQUEST.ATSAuditName# #REQUEST.ATSFindingName#"
			TO="#ToList#" 
			CC="#CCList#" 
			<!--- to="diego.serrano@gensuitellc.com" --->
			FROM="#FromEmail#"
			SUBJECT="#SubjectLine#"
			FOOTER="#Footer#"
			>

		<p>
		<font class="extramediumtext"><i>
		<CFIF ToList EQ "">
			The designated Responsible Person does not have a valid email address, therefore no email was sent.
		<CFELSE>
			<CFOUTPUT>
			An email on this #LCase(REQUEST.ATSAuditName)# #LCase(REQUEST.ATSFindingName)# has been sent to inform the responsible person at <b>#ToList#</b>
			<CFIF CCList IS NOT "">
				(with a cc to <b>#CCList#</b>)
			</CFIF>
			stating this #LCase(REQUEST.ATSFindingName)# has been verified!
			</CFOUTPUT>
		</CFIF>
		</i></font>
	</CFIF>
</CFIF>
<!--- 08/16/02 --->
<!--- 05/28/03 --->



<CFIF Form.Status IS "Closed" AND ((VerifyBy IS NOT "" AND VerifyDate IS "") OR VerifyBy IS NOT VerifyBy_Orig) AND FORM.firstAction IS NOT "Delete">

	<CFQUERY NAME="qGetVerifyEmail" DATASOURCE="#ODBC#">
		SELECT CONTACT_EMAIL
		FROM ltbContact WITH (NOLOCK)
		WHERE Contact_Name = <CF_QUERYPARAM VALUE="#VerifyBy#" CFSQLTYPE="CF_SQL_VARCHAR">
	</CFQUERY>
	<CFQUERY NAME="qGetResponsibleEmail" DATASOURCE="#ODBC#">
		SELECT CONTACT_EMAIL
		FROM ltbContact WITH (NOLOCK)
		WHERE Contact_Name = <CF_QUERYPARAM VALUE="#FORM.ClosePerson#" CFSQLTYPE="CF_SQL_VARCHAR">
	</CFQUERY>
	<CFIF qGetResponsibleEmail.CONTACT_EMAIL IS "">
		<CFQUERY NAME="qGetResponsibleEmail" DATASOURCE="#ODBC#">
			SELECT CONTACT_EMAIL
			FROM ltbContact WITH (NOLOCK)
			WHERE Contact_Name = <CF_QUERYPARAM VALUE="#AccessName#" CFSQLTYPE="CF_SQL_VARCHAR">
		</CFQUERY>
	</CFIF>
	<CFQUERY NAME="qGetResponsibleEmailCC" DATASOURCE="#ODBC#">
		SELECT CONTACT_EMAIL
		FROM ltbContact WITH (NOLOCK)
		WHERE Contact_Name = <CF_QUERYPARAM VALUE="#FORM.ResponPerson#" CFSQLTYPE="CF_SQL_VARCHAR">
	</CFQUERY>

	<CFSET FromEmail = qGetResponsibleEmail.CONTACT_EMAIL>

	<CFSET CCEmail = "">
	<CFIF isDefined("Form.EmailRP") AND Form.EmailRP IS NOT "">
		<CFIF RPEmail.CONTACT_EMAIL IS qGetVerifyEmail.CONTACT_EMAIL>
			<CFSET CCEmail = "">
		<CFELSE>
			<CFSET CCEmail = RPEmail.CONTACT_EMAIL>
			<cfif variables.showCoResponPerson and (form.coResponPerson neq "") and isDefined("CoRPEmail.CONTACT_EMAIL")>
				<CFSET CCEmail = ListAppend(CCEmail, CoRPEmail.CONTACT_EMAIL)>
			</cfif>
		</CFIF>
		<CFIF FORM.MULT_CC IS NOT "">
			<CFSET CCEmail = ListAppend(CCEmail, FORM.MULT_CC)>
		</CFIF>
		<CFSET CCEmail = ListAppend(CCEmail, qGetResponsibleEmailCC.CONTACT_EMAIL)>
	</CFIF>

	<CFIF qGetVerifyEmail.CONTACT_EMAIL IS NOT "">
		<cfif findingTypeLabel eq "">
			<cfset SubjectLine = "#Abr# #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# ##" & #variables.FindingID# & " requires #CLosureVerificationlabelForEmails#: " & OName & ", " & EncodeForHTML(FactoryID) & Loc & " (#GEBusiness#)">
		<cfelse>
	  		<CFSET SubjectLine = "#EncodeForHTML(Form.FindingType)# #Abr# #REQUEST.ATSAuditName# #REQUEST.ATSFindingName# ##" & #variables.FindingID# & " requires #CLosureVerificationlabelForEmails#: " & OName & ", " & EncodeForHTML(FactoryID) & Loc & " (#GEBusiness#)">
	  	</cfif>

	  	<!--- 04/20/04 --->
		<CFSET CCList = ListAppend(CCEmail, FORM.Mult_CC)>
		<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#CLEANEMAIL.cfm"
			ToList = "#qGetVerifyEmail.CONTACT_EMAIL#"
			CCList = "#CCList#"
			Delim = ",">

		<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#CLEANEMAIL.cfm"
			ToList = "#ToList#"
			CCList = "#CCList#"
			Delim = ",">

	  	<!--- "[FindingType] Type Finding Notification" distribution lists setup in contacts permissions if activated for this business   --->
		<cfmodule template="#ContactsMasterPath#distGetList.cfm"
			Application = "#request.ATSAppName#"
		    List = "#Form.FindingType# Type Finding Notification"
			AdditionalToEmails="#ToList#"
			AdditionalCcEmails="#CCList#"
		    SiteID ="#LocID#"
			DSN="#ODBC#">
		<cfset ToList = DistList.ToEmails>
		<cfset CCList = DistList.ccEmails>

	<CFSAVECONTENT VARIABLE="EmailNotes">
		<CFOUTPUT>
		Dear <font color="blue"><i>#Form.VerifyBy#:</b></i><br></FONT>
		<FONT STYLE="color:red;">The #LCase(REQUEST.ATSAuditName)# #LCase(REQUEST.ATSFindingName)# below has been closed and you are designated for its #closureVerificationLabel#.</FONT>
		<BR><BR>
		<b>Please verify the #REQUEST.ATSFindingName# <cfif CLosureVerificationlabelForEmails contains "verification">Closure<cfelse>#CLosureVerificationlabelForEmails#</cfif> information using the link provided below
		<!--- 04/03/03 --->
		<i>
		<CFIF VerifyByDate IS NOT "">
			<STRONG>#VerifyByDatePreposition#</STRONG> #DateFormat(VerifyByDate, 'd-mmm-yy')#
		<CFELSE>
			as soon as possible
		</CFIF>
		</i>
			to avoid delaying final closure of this #LCase(REQUEST.ATSFindingName)#!
		</b>
		</font>
		<CFIF VerifyByDate IS NOT "">
			<BR>
			A verification reminder email will be sent out on #DateFormat(VerifyByDate, 'd-mmm-yy')#.
		</CFIF>
		</CFOUTPUT>
	</CFSAVECONTENT>

	<!--- this email doest show the verifyby row since it is to the verifyby --->
	<CFIF isDefined("iVerifyByLocation") AND iVerifyByLocation NEQ -1>
		<CFSET ArrayDeleteAt(Details, iVerifyByLocation)>
	</CFIF>
	<!--- Closure Verification Needed
	http://cincep09corpge.corporate.ge.com/help/ccenter/ats/Verification_Needed.mht
	--->
	
	<cfif structkeyexists(form,"ContactPerson") && len(trim(form.ContactPerson))>

		<CFQUERY NAME="qGetResponsibleEmailCC" DATASOURCE="#ODBC#">
			SELECT CONTACT_EMAIL
			FROM ltbContact WITH (NOLOCK)
			WHERE Contact_Name = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.ContactPerson#">
		</CFQUERY>
		<CFSET CCList = ListAppend(CCList, qGetResponsibleEmailCC.CONTACT_EMAIL)>

	</cfif>

	<CFMODULE TEMPLATE="#REQUEST.LIBRARY.CUSTOMTAGS.VIRTUALPATH#EmailDetails.cfm"
	 	Details="#Details#"
		DetailsTitle="#REQUEST.ATSFindingName# Details"
	 	Summary="#DetailsSummary#"
		SummaryTitle="#REQUEST.ATSFindingName# Summary"
		ApplicationName="#request.ATSAppName#"
		ApplicationIcon="ats.gif"
		ApplicationURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#audit.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&FactoryID=#EncodeForURL(factoryid)#"
		OrganizationName="#OName#"
		SiteName="#factoryid##Form.Location#"
		BlockTitle="#REQUEST.ATSFindingName# ID##"
		BlockID="#Form.ID#"
		BlockIDURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#status.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&ID=#EncodeForURL(FORM.ID)#&FactoryID=#EncodeForURL(factoryid)#"
		EmailType="#REQUEST.ATSAuditName# #REQUEST.ATSFindingName# Summary"
		EmailNotes="#EmailNotes#"
		StatusTitle="Status" StatusText="#StatusText#"
		LINKURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#audfinding.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&ID=#EncodeForURL(FORM.ID)#&FactoryID=#EncodeForURL(factoryid)#"
		LINKTEXT="Edit this #REQUEST.ATSAuditName# #REQUEST.ATSFindingName#"
		TO="#ToList#" 
		 CC="#CCList#" 
		<!--- to="diego.serrano@gensuitellc.com" --->
		FROM="#FromEmail#"
		SUBJECT="#SubjectLine#"
		MailParam="X-Priority:1,X-MSMail-Priority:High"
		>

		<p>
		<P>
		<FONT CLASS="extramediumtext">
		<I>
		<CFIF ToList EQ "">
			The designated Responsible Person does not have a valid email address, therefore no email was sent.
		<CFELSE>
			<CFOUTPUT>
			An email on this #LCase(REQUEST.ATSAuditName)# #LCase(REQUEST.ATSFindingName)# has been sent to the designated person to verify this
			 #LCase(REQUEST.ATSAuditName)# at <b>#qGetVerifyEmail.CONTACT_EMAIL#</b>
			<CFIF StructKeyExists(qGetResponsibleEmailCC, "CONTACT_EMAIL") && len(trim(qGetResponsibleEmailCC.CONTACT_EMAIL))>
				with cc to the Auditor/Contact Person at <b>#qGetResponsibleEmailCC.CONTACT_EMAIL#</b>
			</CFIF>!
			</CFOUTPUT>
		</CFIF>
		</I>
		</FONT>
	<CFELSE>
		<P>
		<FONT CLASS="extramediumtext">
		<I>
		<CFOUTPUT>
		Sorry, an email on this #LCase(REQUEST.ATSAuditName)# #LCase(REQUEST.ATSFindingName)# could not be sent to the designated person to
		 verify this #LCase(REQUEST.ATSAuditName)# because an email address was not available in the #Abr# Contacts
		  Database for <U>#VerifyBy#</U>!
		</cfoutput>
		</I>
		</FONT>
	</CFIF>
</CFIF>

<CFIF PDAMODE IS true AND FORM.OfflineMode IS "NO">
	<form id="rform" method="POST">
		<input type="hidden" name="qString" value="<cfoutput>#QUERY_STRING#</cfoutput>">
	</form>
	<div class="control-group" style="font-size:13px;">
		<div class="controls">
			<div class="list-item">
				<table class="mHeader" onclick="submitItMobile('#ID#')" style="margin:0px; padding:0px; font-size=13px; border-spacing:0px; border-collapse:collapse;">
					<tr style="cursor: pointer;">
						<th style="text-align:left;"><font style="text-decoration: underline; color:blue;">###ID#</font></th>
						<th style="text-align:right;"><b><cfif status eq "Open" and ClosureDueDate lt now()><font color="red"><cfelseif status eq "Open" and ClosureDueDate gte now()><font color="##FF7519"><cfelse><font color="blue">
						</cfif>#status#</font></b></th>
					</tr>
					</table>
					<table class="detailList" style="margin:0px; cellspacing:0; font-size=13px">
						<tr>
							<td colspan="4"><b>#ClosureDueDateLabel#: </b><cfif ClosureDueDate lt now()>
							<font color="red"><cfelse><font color="black"></cfif>#DateFormat(ClosureDueDate, "dd-mmm-yyyy")#</font></td>
						</tr>
					<tr style="background-color:##E6E6E6">							
						<td colspan="4"><b>Responsible Person:</b> #responperson#</td>
					</tr>
					<tr>
						<td colspan="2"><b>Audit Type:</b><br>#AuditType#</td> 
						<td colspan="2"><b>#REQUEST.ATSFindingName# Type:</b><br>#FindingType#</td>
					</tr>
					<tr style="background-color:##E6E6E6">
						<td colspan="4"><b>#REQUEST.ATSFindingName# Category:</b> #Category#</td>
					</tr>
					<cfif AuditName neq "">
					<tr style="background-color:##E6E6E6">
						<td colspan="4"><b>#REQUEST.ATSAuditName#:</b> #AuditName#</td>
					</tr>
					</cfif>
					<tr>
						<td <cfif (isDefined("qGetSubCOE.SubCOE")) or (BLDG neq "")>colspan="2"<cfelse>colspan="4"</cfif>><b>#centerdeptlabel#:</b> #COE#</td>
						<cfif (isdefined("qGetSubCOE.SubCOE")) or (BLDG neq "")>
						<td colspan="2">
						</cfif>
						<cfif (isdefined("qGetSubCOE.SubCOE"))><b>Sub-Dept:</b> #EncodeForHTML(qGetSubCOE.SubCOE)#</cfif>
						<cfif (isdefined("qGetSubCOE.SubCOE")) and (BLDG neq "")><br></cfif>
						<cfif BLDG neq ""><b>Building:</b> #BLDG#</cfif></td>
					</tr>
					<tr style="background-color:##E6E6E6">
						<td colspan="4"><b>#REQUEST.ATSFindingName# Date:</b> #DateFormat(AuditDate, "dd-mmm-yyyy")#</td>
					</tr>
					<tr>
						<td colspan="4"><b>Description:</b> #EncodeForHTML(Description)#</td>
					</tr>
					<tr style="background-color:##E6E6E6">
						<td colspan="4"><b>Corrective Action:</b> #EncodeForHTML(CorrectiveAction)#</td>
					</tr>
					<tr>
						<td colspan="4">&nbsp;</td>
					</tr>
				</table>
       		</div>
		</div>
	</div>
</CFIF>
</DIV>

<script language="JavaScript">

function submitItMobile(id){
	var rform = document.getElementById('rform');
	<cfoutput>rform.action="status.cfm?SiteID=#LocID#&ID="+id;</cfoutput>
	rform.submit();
	
}

function popupattach(loc,win_no) {
	gsWindowOpen(loc, win_no);
}

</script>
<!--- Insert 3 options --->

<!--- 07/14/03 --->
</CFOUTPUT>
</cfif>

<cfoutput>
	<cfif form.id neq ""><cfset variables.findingid=form.id></cfif>
	<script language="JavaScript">
	function addEditFinding(inForm,sMode) {
		var addLoc = '';
		var RefIDStuff = '';
		var IDStuff = '';
		var Facstuff = '&FactoryID=#FactoryID#';

		if (inForm == null)
		{

			addLoc = window.location.href.split('?')[1];
			addLoc = './audfinding.cfm?' + addLoc;
		}
		else
		{
			addToSite=inForm.Loc.value;
			addToOrg=inForm.Org.value;
			<cfoutput>
				<CFIF Form.RefType NEQ "" AND Form.RefID NEQ "">
					var RefIDStuff = '&RefType=#EncodeForJavascript(Form.RefType)#&RefID=#EncodeForJavascript(Form.RefID)#'
				</CFIF>
				if (sMode=='Edit') {
					var IDStuff='&ID=#id#'
				}
				addLoc='#EncodeForJavascript(sCurrentURL)#audfinding.cfm?Org='+addToOrg+'&Loc='+escape(addToSite)+RefIDStuff+IDStuff+Facstuff;
			</cfoutput>
		}
		window.location=addLoc;
	}


	function copyFinding (inForm,orgid,siteid)
	{
		var copyloc = "";
		copyToSite=inForm.Loc.value;
		copyToOrg=inForm.Org.value;
		if(orgid != null) copyToOrg = orgid;
		if(siteid != null) copyToSite = siteid;

		var scopeString = "?org="+copyToOrg+"&loc="+escape(copyToSite);
		if($.isNumeric(copyToSite,10))
			scopeString = "?siteid="+copyToSite;

		/*[#oname#][#Org#][#Loc#]*/
		//alert(copyToOrg);
		//alert(copyToSite);
		if(copyToOrg=='#org#' && copyToSite=='#Loc#')
		{ /* copy in site - allow user to choose*/
			if('#EncodeForURL(Form.RefType)#'!='' && '#EncodeForURL(Form.RefID)#'!='')
			{
				GS.fn.confirm("This #LCase(REQUEST.ATSFindingName)# has associated #ReferenceLabel# information. Click OK to copy the same #ReferenceLabel# information into the new #LCase(REQUEST.ATSFindingName)#; Cancel to copy the #LCase(REQUEST.ATSFindingName)# without #ReferenceLabel# information!",function(response){
					if(response==true){
						copyloc = '#EncodeForJavascript(sCurrentURL)#audfinding.cfm'+scopeString+'&FindingID=#variables.FindingID#&copyOrg=#EncodeForJavascript(org)#&copyLoc=#EncodeForJavascript(EncodeForURL(Loc))#&CopyRef=Yes';
						window.location=copyloc;
					}else{
						copyloc = '#EncodeForJavascript(sCurrentURL)#audfinding.cfm'+scopeString+'&FindingID=#EncodeForJavascript(variables.FindingID)#&copyOrg=#EncodeForJavascript(org)#&copyLoc=#EncodeForJavascript(EncodeForURL(Loc))#&CopyRef=No&RefType=Audit%20Finding&RefID=#EncodeForJavascript(variables.FindingID)#';
						window.location=copyloc;
					}
				});
				
			}else{
				copyloc = '#EncodeForJavascript(sCurrentURL)#audfinding.cfm'+scopeString+'&FindingID=#variables.FindingID#&CopyRef=No&copyOrg=#EncodeForJavascript(org)#&copyLoc=#EncodeForJavascript(EncodeForURL(Loc))#&RefType=Audit%20Finding&RefID=#variables.FindingID#';
				window.location=copyloc;
			}
		} else { /* copy to other site */
			if('#EncodeForURL(Form.RefType)#'!='' && '#EncodeForURL(Form.RefID)#'!='')
			{
				GS.fn.alert("Note: This #EncodeForJavascript(LCase(REQUEST.ATSFindingName))#'s associated #EncodeForJavascript(ReferenceLabel)# information will NOT be copied into the new #EncodeForJavascript(LCase(REQUEST.ATSFindingName))# since it is being copied to a different #EncodeForJavascript(sitelabel)#.");
			}
			copyloc = '#EncodeForJavascript(sCurrentURL)#audfinding.cfm'+scopeString+'&FindingID=#variables.FindingID#&CopyRef=No&copyOrg=#EncodeForJavascript(org)#&copyLoc=#EncodeForJavascript(EncodeForURL(Loc))#';
			window.location=copyloc;

		}
}


	function popupattach(loc,win_no) {
	    gsWindowOpen(loc, win_no);
	}
	</script>
</cfoutput>
</CFSAVECONTENT>

<CFIF (FORM.PDAMODE IS false OR FORM.OfflineMode IS "YES") AND !StructKeyExists(URL, "replicate") >

<!--- 
	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#Button.cfm" STYLE="display:inline"
	 ONCLICK="addEditFinding(document.addSimForm,'Add');">Add New <cfoutput>#REQUEST.ATSFindingName#</cfoutput></CFMODULE>
	&nbsp;

			<CFIF FORM.firstAction NEQ "Delete">

	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#Button.cfm" STYLE="display:inline"
	 ONCLICK="copyFinding(document.addSimForm);">Copy details to New <cfoutput>#REQUEST.ATSFindingName#</cfoutput></CFMODULE>
	&nbsp;

			</CFIF> --->


	<cfif PDAMode NEQ True><!--- Dot want to do this in PDAMode --->
		<cfif isDefined("AddSimilar")><td colspan=4><cfelse><td colspan=2></td></cfif>
		<CFIF FORM.firstAction NEQ "Delete" and not isdefined("AddSimilar")><td colspan=2></td></CFIF>
		<CFIF FORM.firstAction NEQ "Delete">

					<script>
						function populate(inForm) {
							sendOrg=inForm.Org.value;
							LoadedSites.style.display='none';
							LoadingSites.style.display='inline';
							ClearlocationPassback();
							<cfoutput>crit='audaction_Passback.cfm?odbc=#EncodeForJavascript(odbc)#&loc=#EncodeForJavascript(loc)#&org='+sendOrg</cfoutput>
							LoadlocationPassback(crit)
						}

					</script>
					<cfquery name="OrgQ" datasource="#odbc#">
						select Org.OrgName,[table]
						from Org WITH (NOLOCK)
						INNER JOIN Site WITH (NOLOCK) ON Site.orgname = org.orgname
						WHERE 
							Org.ARCHIVE=0
							AND Site.location is not null
						Order By Org.ORGNAME;
					</cfquery>
	<!--- 
	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#Button.cfm" STYLE="display:inline"
		 ONCLICK="document.getElementById('othersiteselector').style.display='block';populate(document.addSimForm);">Copy to another <cfoutput>#SiteLabel#</cfoutput></CFMODULE> --->
	


					<table cellspacing="0" cellpadding="1" align=center style="display:none;" id="othersiteselector">
						<tr>
							<td>
							<CFOUTPUT>
								<form action="#sCurrentURL#status.cfm" name="addSimForm">
							</CFOUTPUT>
									<fieldset>
									<legend><font class="mediumtext"><b><cfoutput>Select #sitelabel# to copy this #REQUEST.ATSFindingName# to...
									<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#HELPTEXT.cfm">Select the #orglabel#/#sitelabel# and then click "<b>Copy details to New #REQUEST.ATSFindingName#</b>" above.</CFMODULE>

									</cfoutput></b></font></legend>
									<table cellpadding="3" border="0">
										<tr>
											<td>
												<font face="Arial" class="mediumtext"><b><cfoutput>#OrgLabel#</cfoutput>:</b></font><br>
												<table border="0" cellspacing="0" cellpadding="0" style="border: thin inset white;" class="mediumtext">
													<tr>
														<td>
															<select name="Org" onChange="populate(this.form);" class="formfields">

																<cfoutput query="OrgQ">
																	<option value="#Table#" <cfif org eq table>SELECTED</cfif>>#OrgName#
																</cfoutput>
															</select>
														</td>
													</tr>
												</table>
												<font face="Arial" class="mediumtext"><b><cfoutput>#SiteLabel#</cfoutput>:</b></font><br>
												<table border="0" cellspacing="0" cellpadding="0" id="LoadedSites" style="border: thin inset white;" class="mediumtext">
													<tr>
														<td>
															<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#HTMLPASSBACK.cfm" action="HOLDER" PASSBACKNAME="locationPassback" debug="false">
															<select name="Loc" style="color: black;" class="formfields">
																	<cfoutput>
																		<option style="color: black;" value="#Loc#">#Loc#
																	</cfoutput>
															</select>
															</CFMODULE>

														</td>
													</tr>
												</table>
												<table border="0" cellspacing="0" cellpadding="0" id="LoadingSites" style="display:none;" class="extrasmalltext">
													<tr>
														<td>
															Loading...

														</td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
									</fieldset>
									<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#forminclude.cfm">
								</form>
								<!---<script>populate(document.addSimForm);</script>--->
								<cfoutput><font class="smalltext">You must have approved permissions to add a #LCase(REQUEST.ATSFindingName)# for the #sitelabel# selected if different from the default selection.</font></cfoutput>
							</td>
						</tr>
					</table>

		</CFIF>
		<CFIF FORM.firstAction NEQ "Delete" and AuditAttachLoc neq "" and not isdefined("AddSimilar")><td></td></cfif>
	</cfif><!--- Dot run in PDAMode --->

</CFIF>

<CFSET bBlockClose = false>
<CFIF ListFindNoCase(CI_LOCK_ATS_FindingType, Form.FindingType) NEQ 0 AND (ByPassExport IS FALSE OR FORM.EXPCIMandatory contains "CIExport")>
	<CFSET bBlockClose = true>
</CFIF>

<CFSAVECONTENT VARIABLE="sButtons">
	<!--- 04/08/04 --->
	<CFPARAM NAME="RequireClosureVerification" DEFAULT="false">
	<CFIF FORM.firstAction IS NOT "Delete" AND FORM.Status IS NOT "Closed">

		<CFIF IsDefined("Form.ID")><cfset NewID = Form.ID> </CFIF>
		<CFPARAM NAME="FORM.VerifyBy" DEFAULT="">

	<!--- Dont want this for PDAMode Online or PPC Simulated Mode - CJ 05/27/2004 --->
	<!--- <CFIF (PDAMode NEQ True) OR (PDAMode EQ True AND FORM.OfflineMode IS "Yes")>  --->

	<CFPARAM NAME="ShowQuickCloseList" DEFAULT="true">
	<CFIF ShowQuickCloseList IS true AND PDAMode IS NOT true>
	<CFPARAM NAME="OpenSubCA" DEFAULT="0">
	<CFIF (RequireClosureVerification IS true OR FORM.VerifyBy IS NOT "" OR OpenSubCA GT 0) AND bBlockClose IS false>
		<!--- 
		<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#Button.cfm" STYLE="display:block"
			 ONCLICK="window.location = '#sCurrentURL#status.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&FactoryID=#EncodeForURL(FactoryID)#&ID=#FindingID#&AutoClose=yes';">Close this <cfoutput>#LCase(REQUEST.ATSFindingName)#</cfoutput></CFMODULE>
			&nbsp; --->
		<cfoutput>
			<button type="button" class="btn btn-secondary" ONCLICK="window.location = '#sCurrentURL#status.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&FactoryID=#EncodeForURL(FactoryID)#&ID=#FindingID#&AutoClose=yes';">#translator.translate("Mark as Closed")#</button>
		</cfoutput>

	<CFELSE>

		<SCRIPT>
			function submitIt() {
				if (document.Status.cdate.value=='') {
					Ext.MessageBox.alert('Alert', 'Please input a <CFOUTPUT>#EncodeForJavascript(DateClosedLabel)# to Close this #EncodeForJavascript(LCase(REQUEST.ATSFindingName))#</CFOUTPUT>!', function(){document.Status.cdate.focus(); document.getElementById('SubmBtn').disabled=false;});
					return false;
				}
				if (Trim(document.Status.comment.value)=='') {
				<CFOUTPUT>
					Ext.MessageBox.alert('Alert', 'Please input a #EncodeForJavascript(ClosureCommentLabel)# to Close this #EncodeForJavascript(LCase(REQUEST.ATSFindingName))#!', function(){document.Status.comment.focus();document.getElementById('SubmBtn').disabled=false;});
				</CFOUTPUT>

					return false;
				}

				document.Status.submit();
			}
			function Trim(sString) {
				return sString.replace(/^\s+|\s+$/g,'');
			}
		</SCRIPT>
		<STYLE>
		<CFOUTPUT>
			.DIVpopup UL	{ list-style-image: url(#EncodeForJavascript(sCurrentURL)#blue_bullet.gif); margin-left: 18px; margin-top: 2px; margin-bottom: 2px; }
			.DIVpopup LI	{ margin-top: 0px; margin-bottom: 4px; }
		</CFOUTPUT>

			.prompt .divider
			{
				border-top: 1px solid #99BBE8;
			}
			.Prompt .Label
			{
				border-right: 1px solid #99BBE8;
				color: #15428B;
			}
		</STYLE>
		<CFIF bBlockClose IS false>

			<script type="text/javascript">
				function launchCloseWindow(){
				 	$('#closeWindowModal').modal('show');
				 	$('#closeWindowModal').find('.modal-body').html('<div class="loading"></div>');
				 	<cfoutput>
				 		$('##closeWindowModal').find('.modal-body').load('closeFinding.cfm?SiteID=#SiteID#&ID=#EncodeForURL(Form.ID)#&cancelFunction=HideQuickClose#SiteID#')
				 	</cfoutput>
				}
					
			</script>

			<cfoutput>
				<button type="button"
						class="btn btn-default"
						onclick="launchCloseWindow()"
						id="quickCloseButton">
					<i class="fa fa-check" aria-hidden="true"></i> #translator.translate("Mark as Closed")#
				</button>
			</cfoutput>

		</CFIF>
	</CFIF>
</CFIF> <!--- End of (PDAMode NEQ True) OR (PDAMode EQ True AND FORM.OfflineMode IS "Yes")>
--->

</CFIF>

<CFIF FORM.firstAction NEQ "Delete">

	<CFSET sURL = "#sCurrentURL#audfinding.cfm?ID=#variables.FindingID#&Org=#Org#&Loc=#EncodeForURL(Form.Location)#&FactoryID=#EncodeForURL(FactoryID)#">
	<CFIF (#Form.RefType# NEQ "") AND (#Form.RefID# NEQ "")>
		<CFSET sURL = sURL & "&RefType=#EncodeForURL(Form.RefType)#&RefID=#EncodeForURL(Form.RefID)#">
	</CFIF>

</CFIF>

<cfset infotext = EncodeForURL("#REQUEST.ATSFindingName# ID## #variables.FindingID#")>
<CFIF FORM.firstAction NEQ "Delete" and AuditAttachLoc neq "" and not isdefined("AddSimilar")>


<CFIF ListFindNoCase(CI_LOCK_ATS_FindingType, Form.FindingType) NEQ 0 AND CILink.actionsComplete IS "" AND Form.Status neq "Closed">
	<SCRIPT>

		function ExportToCI()
		{
		<cfif isdefined("URL.EXPCI") and URL.EXPCI eq "yes">
			<CFOUTPUT>
				LoadCI('audaction.cfm?siteid=#siteid#&id=#id#&Action=ExportCI&init=#REQUEST.User.AccessID#');
			</CFOUTPUT>
		<cfelse>
			<CFOUTPUT>
				
				GS.fn.confirm("Once exported, this finding will be locked and will not be able to be updated. <BR>Are you sure you want to export this finding to Continuous Improvement?", function(result) {
    				if (result == true) {
        			LoadCI('audaction.cfm?siteid=#siteid#&id=#id#&Action=ExportCI&init=#REQUEST.User.AccessID#');
    				}
				});

			</CFOUTPUT>
		</cfif>
		}
	</SCRIPT>
	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#HTMLPassBack.cfm" PASSBACKNAME="CI" AJAX="true" POSTURL="true"
		SHOWLOADING="true">
		<!--- make the edit button also go away after exporting --->
		<button type="button" class="btn btn-default" ONCLICK="window.location = '<cfoutput>#JSStringFormat(sURL)#</cfoutput>';"><i class="fa fa-pencil" aria-hidden="true"></i> <cfoutput>#translator.translate("Edit")#</cfoutput></button>
		
		<button type="button" class="btn btn-default" ONCLICK="ExportToCI();" id="Export"> Export to Continuous Improvement</button>	
		<cfif isdefined("URL.EXPCI") and URL.EXPCI eq "yes">
			<script>
    			document.getElementById("Export").click();
			</script>
		</cfif>
	</CFMODULE>
<CFELSE>
<!--- 
	<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#Button.cfm" STYLE="display:inline"
	helpwidth="100"
ONCLICK="window.location = '#JSStringFormat(sURL)#';">Edit this <cfoutput>#REQUEST.ATSFindingName#</cfoutput></CFMODULE>
&nbsp; --->

<button type="button" class="btn btn-default" ONCLICK="window.location = '<cfoutput>#JSStringFormat(sURL)#</cfoutput>';"><i class="fa fa-pencil" aria-hidden="true"></i> <cfoutput>#translator.translate("Edit")#</cfoutput></button>

</CFIF>
<button type="button" class="btn btn-default" ONCLICK="copyFinding(document.addSimForm);">
	<i class="fa fa-files-o" aria-hidden="true"></i> <cfoutput>#translator.translate("Copy to New #REQUEST.ATSFindingName#")#</cfoutput>
</button>
<button type="button" id="copyToSiteBtn" class="btn btn-default">
	<cfoutput>#translator.translate("Copy to other #sitelabel#")#</cfoutput>
	<i class="fa fa-share-square-o" aria-hidden="true"></i>
</button>
<cfoutput>
<cfset happydelimiter = chr(8)>
</cfoutput>
</CFIF>

<CFIF FORM.firstAction NEQ "Delete">
	<cfmodule template="#request.library.customtags.virtualpath#extensionshandler.cfm"
	appid="#request.app.appID#" ref_id="#FORM.tblAuditID#" action="audactionpage" siteid="#locId#" audithome="#audithome#" odbc="#odbc#" />
</CFIF>
</DIV>
</CFSAVECONTENT>
	

<CFOUTPUT>#sMessage#</CFOUTPUT>

<CFIF FORM.PDAMODE IS false OR FORM.OfflineMode IS "YES">
	<CFIF Action IS NOT "Delete"><!--- this is not correct for BB/PPC - CJ --->

		<CFSAVECONTENT VARIABLE="sAttach">
			
			<cfmodule TEMPLATE="#Request.Library.CustomTags.VirtualPath#attachDisplay.cfm"
				qAttach="#qGetAttachments#"
				manageIcon="true"
				BusinessID="#BusinessID#"
				AppID="#Request.App.AppID#"
				SiteID="#LocID#"
				RefType="Audit"
				RefID="#FORM.ID#"
				ODBC="#CC_ODBC#"
				PaperClip="false"
				buttonattach="true"
				<!--- customText="<span ID=""qAttach""></span>" --->
				INFOTEXT="#REQUEST.ATSFindingName# ID## #FORM.ID#"/>
			
		</CFSAVECONTENT>
		<!--- doing this so uses can attach additional things with the attachment button at the top --->
		<CFSET Details[ArrayLen(Details)].Value = Details[ArrayLen(Details)].Value & sAttach>
	<!--- 
		<cfdump var="#DetailsSummary#">
			<cfdump var="#Details#"> --->
	
		<!--- <cfdump var="#sButtons#"> --->
	<!--- 
			<CFMODULE TEMPLATE="#REQUEST.LIBRARY.CUSTOMTAGS.VIRTUALPATH#EmailDetails.cfm"
				 	Details="#Details#"
					DetailsTitle="#REQUEST.ATSFindingName# Details"
					DetailsTitleComment="#sButtons#"
				 	Summary="#DetailsSummary#"
					SummaryTitle="#REQUEST.ATSFindingName# Summary"
					ApplicationName="#request.ATSAppName#"cfinclude
					OrganizationName="#OName#"
					SiteName="#factoryid##Form.Location#"
					BlockTitle="#REQUEST.ATSFindingName# ID##"
					BlockID="#Form.ID#"
					BlockIDURL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#status.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&ID=#EncodeForURL(FORM.ID)#&FactoryID=#EncodeForURL(factoryid)#"
					StatusTitle="Status"
					PREVIEW="true"/> --->
	
	<CFELSE>
		<CFOUTPUT>#sButtons#</CFOUTPUT>
	</CFIF>
	<cfoutput>
	
	<cfif action IS NOT "Delete">
		 <cfif not structKeyExists(OrgQ, "table")>
            <cfquery name="OrgQ" datasource="#odbc#">
                select Org.OrgName,[table]
                from Org WITH (NOLOCK) WHERE ARCHIVE=0 and exists (select location from site where orgname = org.orgname)
                Order By Org.ORGNAME;
            </cfquery>
        </cfif>
		<div class="card table-card mb-4">
			<div class="card-header bg-secondary" data-toggle="collapse" href="##details-collapse" role="button">#OName#/#Form.Location#</div>
			<div id="details-collapse" aria-expanded="true" class="collapse show">
				<div class="card-block">
					<table class="table table-hover table-bordered">
						<tbody>
							<tr>
								<th class="bg-active">#translator.translate("#REQUEST.ATSFindingName#")# ID##</th>
								<td><a href="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#status.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&ID=#EncodeForURL(FORM.ID)#&FactoryID=#EncodeForURL(factoryid)#">#form.id#</a></td>
							</tr>
							<cfloop array="#DetailsSummary#" index="idx">
								<tr>
									<th class="bg-active">#translator.translate(idx.title)#</th>
									<td>#idx.value#</td>
								</tr>
							</cfloop>

							<cfloop array="#Details#" index="idx">





								<tr>
									<th class="bg-active">#translator.translate(idx.title)#</th>
									<td>#idx.value#</td>
								</tr>


							</cfloop>
							<cfif structkeyexists(request,'CustomTaggerEnabled') and request.CustomTaggerEnabled eq true and variables.embedCustomTagger eq true>
									<CFMODULE TEMPLATE="#Request.Library.CustomTags.VirtualPath#embedCustomTagger.cfm"  labelclass="" fieldclass="" 
									version="recharged" mode="view" mainDiv="false" useTableElements="true" mainDivClass="">
								</cfif>
							
						
							
						</tbody>
					</table>				
				</div>
			</div>
			<div class="card-footer">#sbuttons#</div>
		</div>
		
		<div class="modal fade" id="closeWindowModal">
		  <div class="modal-dialog" role="document">
		    <div class="modal-content">
		      <div class="modal-header bg-primary">
		        <h5 class="modal-title"><cfoutput>#Translator.Translate("Mark #REQUEST.ATSFindingName# ###findingID# as Closed")#</cfoutput></h5>
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
		          <span aria-hidden="true">&times;</span>
		        </button>
		      </div>
		      <div class="modal-body"></div>
		      <div class="modal-footer">
		      	<button type="button" id="quickCloseSubmit" class="btn btn-secondary"><cfoutput>#translator.translate('Submit')#</cfoutput></button>
		      	<button type="button" class="btn btn-default" data-dismiss="modal"><cfoutput>#translator.translate('Cancel')#</cfoutput></button>
		      </div>
		    </div>
		  </div>
		</div>
		<div class="modal fade" id="copyWindowModal">
		  <div class="modal-dialog" role="document">
		    <div class="modal-content">
		      <div class="modal-header bg-primary">
		        <h5 class="modal-title"><cfoutput>#Translator.Translate("Copy #REQUEST.ATSFindingName# to another #sitelabel#")#</cfoutput></h5>
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
		          <span aria-hidden="true">&times;</span>
		        </button>
		      </div>
		      <div class="modal-body">
		      	<div class="row">
		      		<div class="col-sm-12 col-md-12 col-lg-12 mb-1">
		      			<label class="col-form-label">#OrgLabel# / #SiteLabel#:</label>
		      			<select id="siteCopy" class="form-control gs-simple-select2"></select>		
		      		</div>
		      	</div>
		      </div>
		      <div class="modal-footer">
		      	<button type="button" id="copySubmit" class="btn btn-secondary"><cfoutput>#translator.translate('Submit')#</cfoutput></button>
		      	<button type="button" class="btn btn-default" data-dismiss="modal"><cfoutput>#translator.translate('Cancel')#</cfoutput></button>
		      </div>
		    </div>
		  </div>
		</div>		     
	</cfif>
	</cfoutput>
	<cfmodule 
		template="#request.library.framework.virtualpath#footer.cfm" />
	<script type="text/javascript">
		$(function() {
			$('#quickCloseSubmit').on('click', function(event) {
				event.preventDefault();
				/* Act on the event */
				$('#SubmBtn').trigger('click');
			});
			
			$('#copyToSiteBtn').on('click',  function(event) {
				event.preventDefault();
				$('#copyWindowModal').modal('show');
				/* Act on the event */
			});
			
			$('#copySubmit').on('click', function(event) {
				event.preventDefault();
				if ($('#siteCopy').val()) {
					var orgCopy = $('#siteCopy').val().split('_')[0];
					var siteCopy = $('#siteCopy').val().split('_')[1];
					if (!orgCopy || !siteCopy) {
						GS.fn.alert(<cfoutput>'#Translator.Translate("Please input the #OrgLabel# / #SiteLabel#")#'</cfoutput>);
					} else {
						 $('#copySubmit').prop('disabled', true);
						copyFinding(document.addSimForm, orgCopy, siteCopy);
					}
				} else {
					GS.fn.alert(<cfoutput>'#Translator.Translate("Please input the #OrgLabel# / #SiteLabel#")#'</cfoutput>);
				}
				/* Act on the event */
			});


			$('#siteCopy').select2({
				minimumInputLength: 3
			    , allowClear: false
			    , placeholder:''
			    , tags: []
			    , ajax: {
			    	url: GS.data.paths.domainURL + GS.data.paths.gswPath + "gswcomp.cfm"
			    	, dataType: "json"
			    	, type: "GET"
			    	, delay: 250
			    	, data: function (params) {
			    		return {
			    			gswcomp: 'gsw'
			    			, func: 'searchsitesjson'
			    			, companyid: GS.data.scope.companyID
			    			, busid: GS.data.scope.busID
			    			, query: params.term
			    		}
			    	}
			    	, processResults: function (data) {
			    		var items=[];
			    		$.each(data.data, function(index, val) {
			    			if (GS.data.scope.siteID != val['siteid']) {
				    			items.push({
				    				id: val['orgid'] + '_' + val['siteid']
				    				, text: val['orgname'] + ' / ' + val['sitename']
				    			});
			    			}
			    		});
			    		return {
			    			results: items
			    		};
			    	}
			    }
			});
		});
	</script>
</cfif>


<cfif isDefined("addSimilar")>
	<cfabort>
</cfif>


<!--- End Capturing Data for PDA --->
</CFMODULE>
<CFIF FORM.firstAction IS NOT "delete">
	<cfset watchStatus = "Open">
	<cfset watchListRP = form.responperson>
	<CFSET watchdue = ClosureDueDate>
	<CFSET watchContactNames = form.contactperson>

	<CFSET watchAction = "Open">
	<cfif FORM.firstAction is "add">
		<cfset performedaction = "added">
	<cfelseif FORM.firstAction is "reject">
		<cfset performedaction = "rejected">
	<cfelse> <!--- FORM.firstAction = "edit" --->
		<!--- in an edit you can be just editing the finding, closing the finding, verifying the finding --->
		<CFIF VerifyDate IS NOT "">
			<cfset watchListRP = form.VerifyBy>
			<cfset performedaction = "verified">
			<cfset watchStatus = "Closed">
			<CFSET watchAction = "Verified">
			<!--- the finding was just verififed, call the extension handler to see if they want to do anything --->
			<cfmodule template="#request.library.customtags.virtualpath#extensionshandler.cfm"
				appid="#request.app.appID#" ref_id="#FORM.tblAuditID#" siteid="#locid#" action="verified" audithome="#audithome#" odbc="#odbc#" />
		<CFELSEIF FORM.Status IS "closed">
			<cfset performedaction = "closed">
			<cfset watchStatus = "Closed">
			<CFSET watchAction = "Closed">
			<CFSET watchdue = "">
			<CFIF VerifyBy IS NOT "">
				<cfset watchListRP = form.VerifyBy>
				<!--- If the Responsibility of the watched item has shifted to the verifyby person, keep the RP in the loop by adding them to the contact list --->
				<CFSET watchContactNames = ListAppend(form.contactperson, FORM.ResponPerson)>
				<cfset watchStatus = "Open">
				<CFSET watchAction = "Verification Pending">
				<CFSET watchdue = "">
				<CFIF VerifyByDate IS NOT "">
					<CFSET watchdue = VerifyByDate>
				</CFIF>
			</CFIF>
		<CFELSE>
			<cfset performedaction = "edited">
		</CFIF>
	</cfif>

	<cfset watchListTitle = "#REQUEST.ATSAuditName# #REQUEST.ATSFindingName# ###form.id# - #watchAction#">

	<cfif not isdefined("qGetSiteID")>
		<CFQUERY NAME="qGetSiteID" DATASOURCE="#ODBC#">
			SELECT SiteID
			FROM Site WITH (NOLOCK)
			WHERE OrgName=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
			 AND Location=<CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Loc#">
		</CFQUERY>
	</cfif>

	<cfif FORM.tblAuditID eq "">
		<cfquery name="qGetTblAuditID" datasource="#odbc#">
			select tblAuditID
			from tblAudit with (nolock)
			where orgname = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
				and location = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Form.Location#">
				and id = <CF_QUERYPARAM VALUE="#form.ID#" CFSQLTYPE="CF_SQL_BIGINT">
		</cfquery>
		<cfset FORM.tblAuditID = qGetTblAuditID.tblAuditID>
	</cfif>
	
	<cfif variables.factoryIDHistory>
		<cfset REQUEST.ATSAuditName = factoryid & REQUEST.ATSAuditName>
	</cfif>	
	
	<cfset variables.sReportTitle = "#REQUEST.ATSAuditName# #REQUEST.ATSFindingName#"/>

	<cfmodule 
		template="#request.library.customTags.virtualPath#gsoHistory.cfm"
		type = "action history"
		displaytitle = "#variables.sReportTitle#"
		transTitle = "#Translator.Translate(variables.sReportTitle)#"
		displayAction="#reReplace(performedaction,"(^[a-z]|\s+[a-z])","\U\1","ALL")#"
		displayId="#form.id#"
		link = "status.cfm?Org=#Org#&Loc=#EncodeForURL(Loc)#&ID=#EncodeForURL(form.ID)#"
		siteID = "#qGetSiteID.SiteID#"
		details = ""
		refType = "Audit" <!--- do not change; internal use only, safest setting is a hardcoded value --->
		refID = "#FORM.tblauditid#" />
	<cfset REQUEST.ATSAuditName = variables.globalATSName>
</CFIF>

<cfif openerActionCallback neq ""> <!--- if defined, call opener function with ID and Action --->
	<cfoutput>
		<script>
			try{window.opener.#EncodeForJavascript(openerActionCallback)#('#EncodeForJavascript(form.id)#','#EncodeForJavascript(FORM.firstAction)#');} catch(e){}; // call with ID and Action
		</script>
	</cfoutput>
</cfif>

<cfif FORM.firstAction EQ "add">
<cfparam name="form.appcaller" default="">
<cfparam name="form.issueid" default="">
<cfparam name="form.ergoid" default="">

	<cfif form.AppCaller EQ "Ergo">
		<cfif (form.ergoID NEQ "") AND (form.issueID NEQ "")>
			<CFQUERY NAME="qUpdateRiskRoot" DATASOURCE="#Safety_ODBC#">
				UPDATE ErgoRootCause
				SET RefID = <CF_QUERYPARAM VALUE="#FORM.RefID#" CFSQLTYPE="CF_SQL_VARCHAR">,
				ExportType = 'ats'
				WHERE OrgName = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#OName#">
					AND Location = <CF_QUERYPARAM CFSQLTYPE="CF_SQL_VARCHAR" VALUE="#Loc#">
					AND ErgoID = <CF_QUERYPARAM VALUE="#form.ErgoID#" CFSQLTYPE="CF_SQL_VARCHAR">
					AND IssueID = <CF_QUERYPARAM VALUE="#VAL(form.IssueID)#" CFSQLTYPE="CF_SQL_INTEGER">
			</CFQUERY>
		</cfif>
	</cfif>
</cfif>

<!--- Crazy thing to replicate - It was placed at the very bottom so we wait untill the main finding was completed added, so then we can replicate, not agree? speak now--->
<cfif structKeyExists(form, "selectedSiteIdList") && trim(form.selectedSiteIdList) neq "" && StructKeyExists(FORM, "dateadded") && ABS(dateDiff("s", now(), FORM.dateadded)) LTE 15>
	
	<div class=" container-fluid row">
		<div class="col-md-6">
			<div class="card table-card">
				<div class="card-header bg-secondary" data-toggle="collapse" href="#replicate-collapse" role="button"><cfoutput>#translator.translate("#request.ATSFindingName#s Replicated")#</cfoutput></div>
				<div id="replicate-collapse" aria-expanded="true" class="collapse show">
				<div class="card-block">
					<table class="table table-hover">
						<tbody>
							<cfloop index="iSiteID" list="#form.selectedSiteIdList#" delimiters=",">
								<cfif structKeyExists(form,"replicateRespPerson#iSiteID#")>
									<cfset variables.RRespPerson = Evaluate("form.replicateRespPerson" & iSiteID)>
								<cfelse>
									<cfset variables.RRespPerson = "">
								</cfif>
								
								<cfset variables.replicateContactPerson = structKeyExists(form, "replicateContactPerson#iSiteID#")?form['replicateContactPerson#iSiteID#']:""/>
								<cfset variables.replicateAttachmentFile    = structKeyExists(form, "replicateAttachmentFile")?form['replicateAttachmentFile']:""/>
								<cfif structKeyExists(form, "replicateRespPerson#iSiteID#") and variables.RRespPerson neq "">
									
									<CFSET variables.Response = audit.replicateFinding(siteid="#iSiteID#", tblAuditID="#form.tblAuditID#", ODBC="#ODBC#", AuditHome="#AuditHome#",
																ResponPerson="#form['replicateRespPerson#iSiteID#']#", ContactPerson="#variables.replicateContactPerson#", auditDate="#form['replicateAuditDate#iSiteID#']#",
																closureList="#form['replicateClosureList#iSiteID#']#",  replicateAttachment="#variables.replicateAttachmentFile#", VerifyByDate="", VerifyBy="")>
									<!--- If it was added succesfully, then we need to get its information--->
									<cfif variables.Response.Success eq true>
										<cfset variables.qSite = Lookup.getSite(siteid="#iSiteid#")	>
										<cfset variables.qOrganization = lookup.getSiteOrgLoc(siteid=siteid)/>
										<cfoutput>
										<tr>							
											<th>#variables.qSite.orgname# / #variables.qSite.location#</th>
											<td>
												<b>#request.ATSFindingName# ID##:</b> <a href="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#status.cfm?siteid=#variables.qSite.siteid#&ID=#variables.Response.ID#">#variables.Response.ID#</a>
					    				  		<span class="small">(#translator.translate("click to view/follow-up")#)</span>
					    				  	</td>
										</tr>	
										</cfoutput>
									</cfif>
								</cfif>
							</cfloop>
						</tbody>
					</table>
				</div>
				</div>
			</div>	
		</div>
	</div>
</cfif>

<cfparam name="request.gsAppbetaPreview.betaactive" default="0">
<cfif request.browsercheck.isMobile() && request.gsAppbetaPreview.betaactive && structkeyexists(form,"org") && structkeyexists(form,"location") && form.location neq '' and form.org neq ''>
	<cflocation ADDTOKEN="No" URL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#status.cfm?org=#form.org#&loc=#form.location#&id=#ID#">
</cfif>
<cfif structkeyexists(form,'ats_mobileDeskVer')>
	<cflocation ADDTOKEN="No" URL="#REQUEST.DOMAINPROTOCOL##REQUEST.DOMAINURL##AuditHome#status.cfm?org=#form.orgid#&loc=#form.location#&id=#ID#">
</cfif>
