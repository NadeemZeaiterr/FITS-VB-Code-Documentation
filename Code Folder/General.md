


## Code Snippets By Related Subject Below:

<details>
<summary> <font size="5"> MetaTest Macro to test for a certain property of property page object</font> 
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
