


## Code Snippets By Related Subject Below:

<details>
<summary> <font size="4"> MetaTest Macro to test for a certain property of property page object</font> 
</summary>

```vb 

    Sub TestApply(mgobjToBeTested As MegaObject, mgobjTest As MegaObject, strParameters As String, blnTestResult As Boolean)
      ' Write some code here ...

      ' Return the test result into the blnTestResult variable
      blnTestResult = False

      Dim subj: Set subj = mgobjToBeTested 
      if subj.getProp("~qhH)ueW)Y1nA[Org-Proc - Read Only]") = "Y" then
        blnTestResult = True
      End If 


    End Sub
```
</details>

<details>
<summary> <font size="4"> Check if Two Objects Have Same Class ID</font> 
</summary>

```vb 

  set oTk = oRoot.CurrentEnvironment.Toolkit
  Set oSource = oRoot.GetObjectFromId(Manager.SourceID)
  oTk.SameId(oSource.GetClassId, "~OsUiS9B5iiQ0[Operation]")
```
</details>

<details>
<summary> <font size="4">Check if Current User is a Certain User (Could be Profile)</font> 
</summary>

```vb 

CheckCondition = "Not Head of BU"
	set oTk = oRoot.CurrentEnvironment.Toolkit

	Dim subj: Set subj = mgobjWorkflowSubject
	Dim orgUnit: Set orgUnit = subj.getCollection("~wH8T()duYrA5[Org-Unit-1]").item(1)
	
	Dim headofBU : Set headofBU = orgUnit.getCollection("~pMiU4He)YHDU[Person <System>-1]").item(1)

	If oTk.SameId(headofBU.getID,oRoot.CurrentEnvironment.GetCurrentUserId)then
  	CheckCondition = ""
	End If 

```
</details>

