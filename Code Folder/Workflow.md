## Code Snippets Based on Related Subject Below: 


<details>
<summary> <font size="4"> Update Attributes in Workflow Mapping </font> 
</summary>

```vb 
[MappingSet]
MappingAttribute=~9YMdeRRdInwI[Organizational Process Status]  
Mapping1=~9WMdiSRdIT(I[Organizational Process Status 1 - Update in Progress],~TWMd57RdILdF[Update in Progress] 
Mapping2=~9WMdiSRdIT(I[Organizational Process Status 1 - Update in Progress],~TWMd57RdILdF[Update in Progress]   
```
</details>



<details>
<summary> <font size="4"> Notify a certain Collection only based on Subject (MAcro) </font> 
</summary>

```vb 
    'First Set action to send notificationn and in recepients use :
        ' [@Macro=~Yc3Yinv)Yrs2[Head of BU as Recipient],Function=LocateHeadofBU@]
    ' Head of BU as Recepient Macro Implementation
Option Explicit
Function LocateHeadofBU (workflowStatusInstance As MegaObject) As MegaCollection

	Dim mgRoot : Set mgRoot = workflowStatusInstance.getRoot
	Dim workflowInstanceBase: Set workflowInstanceBase= workflowStatusInstance.getCollection("~txc2pUYqJjKD[Workflow Instance Base]").item(1)
	Dim Wfsubject: Set Wfsubject = workflowInstanceBase.getCollection("~rvc2AGYqJXdC[Subject]").item(1)
	Dim subject: Set subject = mgRoot.getObjectFromId(Wfsubject.getID)

	Dim owner
	Set owner = subject.getCollection("~dh27ejV)YDD3[Org-Process Owner]").item(1)

if owner.exists then 
	Dim company,HeadofBU
	Set company = owner.getCollection("~pGAOnKt)YH09[Company]").item(1)

	if company.exists then
		Set HeadofBU = company.getCollection("~pMiU4He)YHDU[Head of BU]")
		Set LocateHeadofBU = HeadofBU
	End if 

End if 
End Function
```
</details>



<details>
<summary> <font size="4">Launch Wizard from Workflow </font> 
</summary>

```vb 
Option Explicit

Sub SetWizardParameters(Root,objWorkflowSubject)

  Dim objWizard
  Dim subj
  subj = objWorkflowSubject

  Set objWizard = subj.CallMethod("~AfLYxbu47b00[WizardRun]", "~8O0vz2FuYfG9[Connect Deliverables to Gate]")

  objWizard.Run
 
End Sub  

Function ExecuteAction(objWorkflowContextAction, strParameter)

  ' Initializing variables
  Dim mgRoot
  Set mgRoot = objWorkflowContextAction.GetRoot

  Dim mgobjWorkflowSubject
  Set mgobjWorkflowSubject = objWorkflowContextAction.GetWorkflowSubject

  Call SetWizardParameters(mgRoot,mgobjWorkflowSubject)

  ExecuteAction= ""

End Function


```
</details>