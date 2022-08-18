<>-- MetaTest Macro to test for a certain property of property page object:

      Sub TestApply(mgobjToBeTested As MegaObject, mgobjTest As MegaObject, strParameters As String, blnTestResult As Boolean)
      ' Write some code here ...

      ' Return the test result into the blnTestResult variable
      blnTestResult = False

      Dim subj: Set subj = mgobjToBeTested 
      if subj.getProp("~qhH)ueW)Y1nA[Org-Proc - Read Only]") = "Y" then
        blnTestResult = True
      End If 


    End Sub
