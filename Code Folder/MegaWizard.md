

## Code Snippets by Related Suject Below:


<details>
<summary> <font size="4"> Find Current User </font> 
</summary>

```vb 
  	Dim userId: userId = mgRoot.currentEnvironment.getCurrentUserID
	Dim user: Set user = mgRoot.getObjectFromID(userID)
```
</details>





<details>
<summary> <font size="4"> Creating Instances, setting properties and Adding to MetaAssociations </font> 
</summary>

```vb 
Sub OnWizNext(mwctxManager As MegaWizardContext, mppSourcePage As MegaPropertyPage, mppTargetPage As MegaPropertyPage)
Dim mgRoot : Set mgRoot = mwctxManager.Template.getRoot
Dim subj : Set subj =  mwctxManager.Template
Dim types: Set types = subj.getCollection("~rnOVrx72ZnXS[TQM Deliverable Type]")
Dim gates: set gates = subj.getCollection("~UKY9VU2jY1)1[Quality Gate]")
Dim dtype
Dim gate
for each gate in gates
    for each dtype in types
        Dim gateTypeObj : Set gateTypeObj = gate.getCollection("~P4bEsc4jY1HQ[Quality Gate Type]").item(1)
        Dim gateType
        gateType = gateTypeObj.getProp("~dP0vQkDuYLQ8[Gate - Type]")
        Dim deliverableRelatedGate
        deliverableRelatedGate = dtype.getProp("~BnFrRHEwYrLb[Related Gate]")
            Dim dict: Set dict = CreateObject("Scripting.Dictionary")
            if dtype.getProp("~emOVLz72ZvcS[Finance Team]") = "Y" then
                dict.Add "Finance", 1
            end if
            if dtype.getProp("~0mOVEz72ZraS[HR Team]") = "Y" then
                dict.Add "HR", 1
            end if
            if dtype.getProp("~xnOVkz72ZXgS[IT Team]") = "Y" then
                dict.Add "IT", 1
            end if
        dict.Add "", 1
Dim varKey
        For each varKey In dict.Keys()
        if deliverableRelatedGate = gateType then
            Dim ic : Set ic = mgRoot.getCollection("~a7bEFk3jYXHO[TQM Deliverable]").CallFunction("~GuX91iYt3z70[InstanceCreator]")
            ic.Property("~o1flDlBwY5Qi[Status - TQM Deliverable]") = "NS"
                ic.Property("~jhweMhCzY5MI[Connected Gate Name]") = gateTypeObj.getProp("~Z20000000D60[Short Name]")
                'ic.Property("~vePGq8auY189[Deliverable - End Date]") = dtype.getProp("~vePGq8auY189[Deliverable - End Date]")
                'ic.Property("~g02FtvZuYnOI[Deliverable - Start Date]") = dtype.getProp("~g02FtvZuYnOI[Deliverable - Start Date]")
                ic.Property("~XKtEBIaxY9DW[Code - TQM Deliverable]") = dtype.getProp("~XKtEBIaxY9DW[Code - TQM Deliverable]")
                ic.Property("~Fc4SyTQkYX5G[Deliverable Evaluation Status]") = "M"
                ic.Property("~f10000000b20[Comment]") = dtype.getProp("~f10000000b20[Comment]")
                ic.Property("~WfP1E(XwYn5R[Max Number of Major Non-Conformance] ") = CInt(dtype.getProp("~WfP1E(XwYn5R[Max Number of Major Non-Conformance]"))
                ic.Property("~GgP1S(XwYv7R[Max Number of Minor Non-Conformance] ") = CInt(dtype.getProp("~GgP1S(XwYv7R[Max Number of Minor Non-Conformance]"))
                ic.Property("~KgP1uMXwYPdP[Approach - Deliverable]") = dtype.getProp("~KgP1uMXwYPdP[Approach - Deliverable]")
                ic.Property("~bBwSQHywY9vq[Nature - TQM Deliverable]") = dtype.getProp("~bBwSQHywY9vq[Nature - TQM Deliverable]")
                ic.Property("~kf2kvaCyYrzM[Metrics - TQM Deliverable]") = dtype.getProp("~kf2kvaCyYrzM[Metrics - TQM Deliverable]")
                    ic.Property("~210000000900[Name]") = subj.getProp("~Z20000000D60[Short Name]")& " - "& dType.getProp("~Z20000000D60[Short Name]") &" " & varKey
            Dim id
            id = ic.Create
            'Dim delType: set delType = mgRoot.getSelection("Select [TQM Deliverable Type] Where [Related Gate] = '"& gateType &"'").item(1)
            Dim deliverable : Set deliverable= mgRoot.GetObjectFromId(id)
            Dim col: Set col = gate.getCollection("~w7bEkm3jYvOO[TQM Deliverable]").Add(deliverable)
            Dim delTypeCol : Set delTypeCol = deliverable.GetCollection("~E5bEHn3jYzRO[TQM Deliverable Type]").Add(dtype)
                     Dim QAactvs: Set QAactvs = dtype.getCollection("~sf35HsJyY1wP[Quality Assurance]")
                     Dim QCactvs: Set QCactvs = dtype.getCollection("~)g35esJyY9zP[Quality Control]")
                 Dim actv1, actv2
                For each actv1 in QAactvs
                        Dim Dactvs : Set Dactvs = deliverable.getCollection("~56bEHH4jYfUP[Quality Assurance]").Add(actv1)
                    Next
        For each actv2 in QCactvs
                Dim Dactvs2 : Set Dactvs2 = deliverable.getCollection("~k7bE8I4jYjXP[Quality Control]").Add(actv2)
            Next
        End if
    Next
    Next
Next
End Sub
```
</details>


<details>
<summary> <font size="4"> Deleting All Objects of a MetaAssociation </font> 
</summary>

```vb 
Dim msg 
	msg = "Attention: Going back will reset your work on this page!"
	Call MsgBox(msg,64,"Attention!")
Dim subj : Set subj = mwctxManager.Template
Dim gates: Set gates = subj.getCollection("~UKY9VU2jY1)1[Quality Gate]")

Dim gate

For each gate in gates 
		Dim deliverables: set deliverables = gate.getCollection("~w7bEkm3jYvOO[TQM Deliverable]")
		Dim del 
	while deliverables.count <> 0 
		For each del in deliverables 
				del.delete()
		Next 
	wend
Next 

```
</details>