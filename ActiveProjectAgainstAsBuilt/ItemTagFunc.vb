Option Strict Off
Option Explicit On

Imports System.Xml
Imports Llama

<System.Runtime.InteropServices.ProgId("ItemTag.ItemTagFunc")> Public Class ItemTagFunc
    Implements LMForeignCalc.ILMForeignCalc
    '*************************************************************************
    'Copyright © 1998, Intergraph Corporation. All rights reserved.
    '
    'File
    '     ItemTagFunc.Cls
    '
    'Author
    '     Raghu Veeramreddy
    '
    'Abstract
    '     Implements ILMForeignCalc class.
    '
    'Description
    '     Validates and calculates item tags.
    '
    'Modifications:
    'Date:          Changes:
    '5/10/00        Added standard document header; removed obsolete code -- dmw.
    '6/26/00        Added error handling -- dmw
    '*************************************************************************


    'do not change the values for these constants
    Const CONST_EquipmentPlantItemIndex As Short = 16 'I18_OK
    Const CONST_RoomPlantItemIndex As Short = 171 'I18_OK
    Const CONST_InstrumentPlantItemIndex As Short = 5 'I18_OK
    Const CONST_InstrumentPlantItemIndex1 As Short = 41 'I18_OK
    Const CONST_NozzlePlantItemIndex As Short = 21 'I18_OK
    Const CONST_PipeRunPlantItemIndex As Short = 4 'I18_OK
    Const CONST_DuctRunPlantItemIndex As Short = 168 'I18_OK
    Const CONST_InstrLoopPlantItemIndex As Short = 46 'I18_OK
    Const CONST_EquipmentCompItemIndex As Short = 6 'I18_OK
    Const CONST_PipingCompItemIndex As Short = 35 'I18_OK
    Const CONST_FalseIndex As Short = 1 'I18_OK
    Const CONST_TrueIndex As Short = 2 'I18_OK
    Const CONST_EquipmentItemName As String = "Equipment" 'I18_OK
    Const CONST_RoomItemName As String = "Room" 'I18_OK
    Const CONST_InstrumentItemName As String = "Instrument" 'I18_OK
    Const CONST_NozzleItemName As String = "Nozzle" 'I18_OK
    Const CONST_PipeRunItemName As String = "PipeRun" 'I18_OK
    Const CONST_DuctRunItemName As String = "DuctRun" 'I18_OK
    Const CONST_InstrLoopItemName As String = "InstrLoop" 'I18_OK
    Const CONST_PlantItemTypeAttributeName As String = "PlantItemType" 'I18_OK
    Const CONST_ItemTagAttributeName As String = "ItemTag" 'I18_OK
    Const CONST_TagSequenceNoAttributeName As String = "TagSequenceNo" 'I18_OK
    Const CONST_TagSuffixAttributeName As String = "TagSuffix" 'I18_OK
    Const CONST_TagPrefixAttributeName As String = "TagPrefix" 'I18_OK
    Const CONST_MeasuredVariableCodeAttributeName As String = "MeasuredVariableCode" 'I18_OK
    Const CONST_InstrumentTypeModifierAttributeName As String = "InstrumentTypeModifier" 'I18_OK
    Const CONST_LoopTagSuffixAttributeName As String = "LoopTagSuffix" 'I18_OK
    Const CONST_OptionSettingAttributeName As String = "OptionSetting" 'I18_OK
    Const CONST_NameAttributeName As String = "Name" 'I18_OK

    Const CONST_EquipNextSeqNoAttributeName As String = "EquipNextSeqNo" 'I18_OK
    Const CONST_RoomNextSeqNoAttributeName As String = "Room Next Sequence Number" 'I18_OK

    Const CONST_PipeRunNextSeqNoAttributeName As String = "PipeRunNextSeqNo" 'I18_OK
    Const CONST_DuctRunNextSeqNoAttributeName As String = "Duct Run Next Sequence Number" 'I18_OK
    Const CONST_InstrNextSeqNoAttributeName As String = "InstrNextSeqNo" 'I18_OK
    Const CONST_InstrLoopNextSeqNoAttributeName As String = "InstrLoopNextSeqNo" 'I18_OK
    Const CONST_OperFluidCodeAttributeName As String = "OperFluidCode" 'I18_OK
    Const CONST_PartofPlantItem_SP_IDAttributeName As String = "PartofPlantItem.SP_ID" 'I18_OK
    Const CONST_ClassAttributeName As String = "Class" 'I18_OK
    Const CONST_LoopFunctionAttributeName As String = "LoopFunction" 'I18_OK

    Const CONST_SP_EquipmentIDAttributeName As String = "SP_EquipmentID" 'I18_OK
    Const CONST_SP_RoomIDAttributeName As String = "SP_RoomID" 'I18_OK

    Const CONST_ItemStatus As String = "ItemStatus" 'I18_OK
    Const Const_ItemStatusValue As String = "1" 'I18_OK
    Const Const_SignalRunItemName As String = "SignalRun" 'I18_OK
    Const CONST_PartOfIDAttributeName As String = "SP_PartOfID" 'I18_OK
    Const CONST_IsBulkItemAttributeName As String = "IsBulkItem" 'I18_OK
    Const CONST_SP_IDAttributeName As String = "SP_ID" 'I18_OK

    'New assign loop automaticly parameters
    Const CONST_CatalogExplorerrootpath As String = "Catalog Explorer root path"
    Private Const Const_ItemType As String = "ITEMTYPE"
    Private Const Const_Entity As String = "Entity"
    Const CONST_MatchAll As Integer = -1
    Const CONST_ItemTagAttribute As String = "ItemTag"
    Const CONST_ModelItemIDAttribute As String = "SP_ModelItemID"
    Const CONST_ItemStatusAttribute As String = "ItemStatus"
    Const CONST_PlantItemGroupItemName As String = "PlantItemGroup"

    Private m_isUIEnabled As Boolean
    Private m_lngID As String
    Private m_NozzleSeqNo As Integer
    Private m_PrevNozzleEqID As String
    Private m_objLMNozzles As Llama.LMNozzles
    Private m_DuplicateTagCheckScope As eDuplicateTagCheckScope
    Private colMultiAssign As Collection

    Private Enum eDuplicateTagCheckScope
        ActivePlant = 0 'this is the original check
        ActiveProjAgainstAsBuilt = 1 'this means the local project and asbuilt
        ActiveProjAgainstAsBuiltAndProjs = 2 'this means asbuilt and all its projects
    End Enum

    Private Function GetChildNozzles(ByRef objLMPlantItem As Llama.LMPlantItem) As Llama.LMAItems
        Dim objChildPlantItem As Llama.LMPlantItem
        Dim objChildPlantItems As Llama.LMPlantItems
        Dim objLMNozzle As Llama.LMNozzle

        objChildPlantItems = objLMPlantItem.ChildPlantItemPlantItems

        For Each objChildPlantItem In objChildPlantItems
            objChildPlantItem.Attributes.BuildAttributesOnDemand = True

            If (objChildPlantItem.ItemTypeName = CONST_NozzleItemName) Then

                objLMNozzle = objChildPlantItem.DataSource.GetNozzle(CStr(objChildPlantItem.Id))
                m_objLMNozzles.Add(objLMNozzle.AsLMAItem)
            Else
                Call GetChildNozzles(objChildPlantItem)
            End If
        Next objChildPlantItem

        objChildPlantItem = Nothing
        objChildPlantItems = Nothing
        objLMNozzle = Nothing
        GetChildNozzles = Nothing
    End Function

    Private Sub Class_Initialize_Renamed()
        m_objLMNozzles = Nothing
        'Default checks only the active plant
        m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActivePlant
    End Sub
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub

    Private Sub Class_Terminate_Renamed()
        m_objLMNozzles = Nothing
    End Sub
    Protected Overrides Sub Finalize()
        Class_Terminate_Renamed()
        MyBase.Finalize()
    End Sub


    Private Function ILMForeignCalc_DoCalculate(ByRef Datasource As Llama.LMADataSource, ByRef Items As Llama._LMAItems, ByRef strPropertyName As String, ByRef varValue As Object) As Boolean Implements LMForeignCalc.ILMForeignCalc.DoCalculate
        Return False
    End Function

    Private Function ILMForeignCalc_DoValidateItem(ByRef Datasource As Llama.LMADataSource, ByRef Items As Llama._LMAItems, ByRef Context As LMForeignCalc.ENUM_LMAValidateContext) As Boolean Implements LMForeignCalc.ILMForeignCalc.DoValidateItem
    End Function

    Private Function ILMForeignCalc_DoValidateProperty(ByRef Datasource As Llama.LMADataSource, ByRef Items As Llama._LMAItems, ByRef strPropertyName As String, ByRef varValue As Object) As Boolean Implements LMForeignCalc.ILMForeignCalc.DoValidateProperty
        m_isUIEnabled = True
        ILMForeignCalc_DoValidateProperty = UpdateItemTag(Datasource, Items, varValue, strPropertyName)
    End Function

    Private Sub ILMForeignCalc_DoValidatePropertyNoUI(ByRef Datasource As Llama.LMADataSource, ByRef Items As Llama._LMAItems, ByRef strPropertyName As String, ByRef varValue As Object) Implements LMForeignCalc.ILMForeignCalc.DoValidatePropertyNoUI
        m_isUIEnabled = False
        UpdateItemTag(Datasource, Items, varValue, strPropertyName)
    End Sub

    Private Function UpdateItemTag(ByRef Datasource As Llama.LMADataSource, ByRef Items As Llama.LMAItems, ByRef varValue As Object, ByRef strPropertyName As String) As Boolean

        Dim objLMAItem As Llama.LMAItem
        Dim BNotAnyCase As Boolean


        On Error GoTo ErrorHandler
        UpdateItemTag = False
        'm_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActivePlant
        m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuilt
        'm_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuiltAndProjs

        'If the workshare type is a satellite then only active plant checking is supported
        If Datasource.IsSatellite = True Then
            m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActivePlant
        End If

        colMultiAssign = Nothing
        colMultiAssign = New Collection

        BNotAnyCase = True
        For Each objLMAItem In Items
            objLMAItem.Attributes.BuildAttributesOnDemand = True

            Select Case objLMAItem.Attributes(CONST_PlantItemTypeAttributeName).Index

                Case CONST_PipeRunPlantItemIndex ' 4
                    'Update the PipeRun ItemTag
                    UpdateItemTag = UpdatePipeRunTag(Datasource, Datasource.GetPipeRun((objLMAItem.Id)), varValue, strPropertyName)
                Case CONST_DuctRunPlantItemIndex ' 168
                    'Update the DuctRun ItemTag
                    UpdateItemTag = UpdateDuctRunTag(Datasource, Datasource.GetDuctRun((objLMAItem.Id)), varValue, strPropertyName)
                Case CONST_InstrumentPlantItemIndex, CONST_InstrumentPlantItemIndex1
                    'Update the Instrument ItemTag
                    objLMAItem.Attributes.BuildAttributesOnDemand = True
                    If objLMAItem.Attributes.Item("ItemTypeName").Value = "Instrument" Then
                        UpdateItemTag = UpdateInstrTag(Datasource, Datasource.GetInstrument((objLMAItem.Id)), varValue, strPropertyName)
                    ElseIf objLMAItem.Attributes.Item("ItemTypeName").Value = "SignalRun" Then
                        UpdateItemTag = UpdateSignalRunTag(Datasource, Datasource.GetSignalRun((objLMAItem.Id)), varValue, strPropertyName)
                    End If
                Case CONST_EquipmentPlantItemIndex
                    'Update the Equipment ItemTag
                    If objLMAItem.Attributes(CONST_ClassAttributeName).Index = CONST_EquipmentCompItemIndex Then
                        UpdateItemTag = True
                        BNotAnyCase = False
                    Else
                        UpdateItemTag = UpdateEquipTag(Datasource, Datasource.GetEquipment((objLMAItem.Id)), varValue, strPropertyName)
                    End If
                Case CONST_RoomPlantItemIndex
                    UpdateItemTag = UpdateRoomTag(Datasource, Datasource.GetRoom((objLMAItem.Id)), varValue, strPropertyName)
                Case CONST_NozzlePlantItemIndex
                    'Updates the Nozzle ItemTag
                    UpdateItemTag = UpdateNozzleTag(Datasource, Datasource.GetNozzle((objLMAItem.Id)), varValue, strPropertyName)
                    BNotAnyCase = True
                Case CONST_InstrLoopPlantItemIndex
                    If objLMAItem.Attributes("PlantItemGroupType").Index = 6 Then
                        UpdateItemTag = UpdateLoopTag(Datasource, Datasource.GetInstrLoop((objLMAItem.Id)), varValue, strPropertyName)
                    Else
                        UpdateItemTag = True
                    End If
                Case CONST_PipingCompItemIndex
                    'Update the PipingComponent ItemTag
                    UpdateItemTag = True
                    BNotAnyCase = False
                Case Else
                    UpdateItemTag = True
            End Select

            'safety -- do not remove
            If BNotAnyCase = True Then
                If strPropertyName = CONST_TagSequenceNoAttributeName And Items.Count < 2 Then

                    If Not IsDBNull(objLMAItem.Attributes(CONST_TagSequenceNoAttributeName).Value) And objLMAItem.Attributes(CONST_PlantItemTypeAttributeName).Index <> CONST_EquipmentCompItemIndex Then

                        varValue = objLMAItem.Attributes(CONST_TagSequenceNoAttributeName).Value

                    End If

                End If
            End If

            If UpdateItemTag = False Then
                varValue = ""
                Exit For
            End If
        Next objLMAItem

        'This is something We have to look into
        If BNotAnyCase = True Then
            If Items.Count > 1 And strPropertyName <> CONST_OperFluidCodeAttributeName Then
                UpdateItemTag = False
                Exit Function
            ElseIf Items.Count > 1 And strPropertyName = CONST_OperFluidCodeAttributeName Then
                UpdateItemTag = True
            ElseIf Items.Count = 1 Then
                UpdateItemTag = UpdateItemTag
            End If
        Else
            UpdateItemTag = True
        End If

        Exit Function

ErrorHandler:

        LogError("ItemTag - ItemTagFunc::UpdateItemTag --> " & Err.Description)
        If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuilt Then
            LogError("ItemTag - ItemTagFunc::UpdateItemTag --> Duplicate Tag Scope: Active Project Against AsBuilt")
        ElseIf m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuiltAndProjs Then
            LogError("ItemTag - ItemTagFunc::UpdateItemTag --> Duplicate Tag Scope: Active Project Against AsBuilt and its projects")
        Else
            LogError("ItemTag - ItemTagFunc::UpdateItemTag --> Duplicate Tag Scope: Active Project or Plant")
        End If

    End Function

    Private Function UpdateInstrTag(ByRef Datasource As Llama.LMADataSource, ByRef Item As Llama.LMInstrument, ByRef varValue As Object, ByRef strPropertyName As String) As Boolean

        Dim bUniqueStatus As Boolean
        Dim colTagValues As Collection
        Dim objFilter As Llama.LMAFilter
        Dim strDupMessage As String
        Dim strItemTag As String
        Dim strLocTagSeqNo As String
        Dim blnUnique As Boolean
        Dim intResponse As Short
        Dim sExist As String
        Dim objLMInstrLoop As LMInstrLoop = Nothing
        Dim iCount As Long
        Dim bHasLoop As Boolean

        On Error Resume Next

        Item.Attributes.BuildAttributesOnDemand = True
        strDupMessage = ""

        'Use the collections unique index feature to avoid a bunch of if statements to determine what property was passed in
        colTagValues = New Collection

        With colTagValues
            .Add(Trim(VariantToString(varValue)), strPropertyName)

            Select Case strPropertyName

                Case CONST_TagSuffixAttributeName
                    .Add(Trim(VariantToString((Item.MeasuredVariableCode))), CONST_MeasuredVariableCodeAttributeName)
                    .Add(Trim(VariantToString((Item.LoopTagSuffix))), CONST_LoopTagSuffixAttributeName)
                    .Add(Trim(VariantToString((Item.InstrumentTypeModifier))), CONST_InstrumentTypeModifierAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)

                Case CONST_TagSequenceNoAttributeName
                    .Add(Trim(VariantToString((Item.MeasuredVariableCode))), CONST_MeasuredVariableCodeAttributeName)
                    .Add(Trim(VariantToString((Item.LoopTagSuffix))), CONST_LoopTagSuffixAttributeName)
                    .Add(Trim(VariantToString((Item.InstrumentTypeModifier))), CONST_InstrumentTypeModifierAttributeName)
                    .Add(Trim(VariantToString((Item.TagSuffix))), CONST_TagSuffixAttributeName)

                Case CONST_MeasuredVariableCodeAttributeName
                    .Add(Trim(VariantToString((Item.LoopTagSuffix))), CONST_LoopTagSuffixAttributeName)
                    .Add(Trim(VariantToString((Item.InstrumentTypeModifier))), CONST_InstrumentTypeModifierAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)
                    .Add(Trim(VariantToString((Item.TagSuffix))), CONST_TagSuffixAttributeName)

                Case CONST_LoopTagSuffixAttributeName
                    .Add(Trim(VariantToString((Item.MeasuredVariableCode))), CONST_MeasuredVariableCodeAttributeName)
                    .Add(Trim(VariantToString((Item.LoopTagSuffix))), CONST_LoopTagSuffixAttributeName)
                    .Add(Trim(VariantToString((Item.InstrumentTypeModifier))), CONST_InstrumentTypeModifierAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)
                    .Add(Trim(VariantToString((Item.TagSuffix))), CONST_TagSuffixAttributeName)

                Case CONST_InstrumentTypeModifierAttributeName
                    .Add(Trim(VariantToString((Item.MeasuredVariableCode))), CONST_MeasuredVariableCodeAttributeName)
                    .Add(Trim(VariantToString((Item.LoopTagSuffix))), CONST_LoopTagSuffixAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)
                    .Add(Trim(VariantToString((Item.TagSuffix))), CONST_TagSuffixAttributeName)

                Case Else
                    .Add(Trim(VariantToString((Item.MeasuredVariableCode))), CONST_MeasuredVariableCodeAttributeName)
                    .Add(Trim(VariantToString((Item.LoopTagSuffix))), CONST_LoopTagSuffixAttributeName)
                    .Add(Trim(VariantToString((Item.InstrumentTypeModifier))), CONST_InstrumentTypeModifierAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)
                    .Add(Trim(VariantToString((Item.TagSuffix))), CONST_TagSuffixAttributeName)
            End Select

        End With

        strLocTagSeqNo = colTagValues.Item(CONST_TagSequenceNoAttributeName)

        'need to update loop relation only if this two properties are involved
        If strPropertyName = CONST_MeasuredVariableCodeAttributeName Or strPropertyName = CONST_TagSequenceNoAttributeName Then
            'need to update loop relation only if this two properties have been changed
            If Item.MeasuredVariableCode <> colTagValues(CONST_MeasuredVariableCodeAttributeName) Or Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value <> strLocTagSeqNo Or IsDBNull(Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value) Then
                'need to update loop relation only if this two properties have values
                If Len(colTagValues(CONST_MeasuredVariableCodeAttributeName)) > 0 And Len(strLocTagSeqNo) > 0 Then
                    'find loop to auto-relate
                    objLMInstrLoop = GetMatchingLoopTag(Datasource, Item, colTagValues(CONST_MeasuredVariableCodeAttributeName), strLocTagSeqNo)

                    If Not objLMInstrLoop Is Nothing Then
                        bHasLoop = True

                        'Need to update tag suffix for the itemtag generate
                        colTagValues.Remove(CONST_LoopTagSuffixAttributeName)
                        colTagValues.Add(Trim$(VariantToString(objLMInstrLoop.Attributes(CONST_TagSuffixAttributeName).Value)), CONST_LoopTagSuffixAttributeName)
                    End If
                End If
            End If
        End If

        strItemTag = BuildInstrumentTag(colTagValues.Item(CONST_MeasuredVariableCodeAttributeName), colTagValues.Item(CONST_InstrumentTypeModifierAttributeName), strLocTagSeqNo, colTagValues.Item(CONST_LoopTagSuffixAttributeName), colTagValues.Item(CONST_TagSuffixAttributeName))

        UpdateInstrTag = True

        If Len(strLocTagSeqNo) > 0 Then
            'Check for uniqueness
            bUniqueStatus = CheckForInstrUnique(Datasource, Item, strItemTag, strDupMessage)
            If bUniqueStatus Then
                bUniqueStatus = Datasource.CheckUniqueInCache(CONST_InstrumentItemName, strItemTag, (Item.Id))
            End If

            If bUniqueStatus = False Then
                For iCount = 1 To Item.PlantItemGroups.Count()
                    If Item.PlantItemGroups.Nth(iCount).Attributes("PlantItemGroupType").Index = 6 Then
                        bHasLoop = True
                        Exit For
                    End If
                Next

                If bHasLoop Then 'we have a loop
                    If m_isUIEnabled Then
                        intResponse = MsgBox(My.Resources.str5003, MsgBoxStyle.OkOnly + MsgBoxStyle.SystemModal + MsgBoxStyle.Exclamation, My.Resources.str5001)
                    End If
                    intResponse = MsgBoxResult.No
                Else
                    If m_isUIEnabled Then
                        intResponse = MsgBox(strDupMessage & " " & My.Resources.str5009, MsgBoxStyle.YesNo + MsgBoxStyle.SystemModal + MsgBoxStyle.Exclamation, My.Resources.str5001)
                    Else
                        intResponse = MsgBoxResult.Yes
                    End If
                End If

                If intResponse = MsgBoxResult.Yes Then
                    blnUnique = False

                    Do While blnUnique = False

                        'Get the next Available tag sequence number from Options Manager
                        strLocTagSeqNo = GetNextAvailTagSeqNo(Datasource, CONST_NameAttributeName, CONST_InstrNextSeqNoAttributeName)

                        strItemTag = BuildInstrumentTag(colTagValues.Item(CONST_MeasuredVariableCodeAttributeName), colTagValues.Item(CONST_InstrumentTypeModifierAttributeName), strLocTagSeqNo, "", colTagValues.Item(CONST_TagSuffixAttributeName))

                        blnUnique = CheckForInstrUnique(Datasource, Item, strItemTag, strDupMessage)
                    Loop
                Else
                    'roll back all data.
                    varValue = Item.Attributes.Item(strPropertyName).Value
                    UpdateInstrTag = False
                End If
            End If
        End If

        If bUniqueStatus = False Then
            Item.LoopTagSuffix = ""
            If Item.Id <> m_lngID Then
                m_lngID = Item.Id
            End If
        End If

        'Update the values
        If UpdateInstrTag Then

            With Item
                If Not objLMInstrLoop Is Nothing Then
                    'relate Instrument-Loop
                    .PlantItemGroups.Add(objLMInstrLoop.AsLMPlantItemGroup.AsLMAItem)

                    objLMInstrLoop = Nothing
                End If

                If Len(strLocTagSeqNo) > 0 Or strPropertyName = CONST_TagSequenceNoAttributeName Then
                    .Attributes.Item(CONST_TagSequenceNoAttributeName).Value = strLocTagSeqNo
                End If

                If strLocTagSeqNo = "" Then
                    .ItemTag = System.DBNull.Value
                Else
                    If Len(colTagValues.Item(CONST_MeasuredVariableCodeAttributeName)) > 0 Then
                        .MeasuredVariableCode = colTagValues.Item(CONST_MeasuredVariableCodeAttributeName)
                    End If

                    If Len(colTagValues.Item(CONST_TagSuffixAttributeName)) > 0 Then
                        .TagSuffix = colTagValues.Item(CONST_TagSuffixAttributeName)
                    End If

                    If Len(colTagValues.Item(CONST_LoopTagSuffixAttributeName)) > 0 And bUniqueStatus Then
                        .LoopTagSuffix = colTagValues.Item(CONST_LoopTagSuffixAttributeName)
                    End If

                    If Len(colTagValues.Item(CONST_InstrumentTypeModifierAttributeName)) > 0 Then
                        .InstrumentTypeModifier = colTagValues.Item(CONST_InstrumentTypeModifierAttributeName)
                    End If

                    If Len(strItemTag) > 0 Then
                        .ItemTag = strItemTag
                    End If

                End If

                If strLocTagSeqNo = "" Then
                    If strPropertyName = CONST_InstrumentTypeModifierAttributeName Then
                        If Len(colTagValues.Item(CONST_InstrumentTypeModifierAttributeName)) > 0 Then
                            .InstrumentTypeModifier = colTagValues.Item(CONST_InstrumentTypeModifierAttributeName)
                        End If
                    ElseIf strPropertyName = CONST_MeasuredVariableCodeAttributeName Then
                        If Len(colTagValues.Item(CONST_InstrumentTypeModifierAttributeName)) > 0 Then
                            .InstrumentTypeModifier = colTagValues.Item(CONST_InstrumentTypeModifierAttributeName)
                        End If
                        If Len(colTagValues.Item(CONST_MeasuredVariableCodeAttributeName)) > 0 Then
                            .Attributes.Item(CONST_MeasuredVariableCodeAttributeName).Value = colTagValues.Item(CONST_MeasuredVariableCodeAttributeName)
                        End If
                    End If
                    If strPropertyName = CONST_TagSuffixAttributeName Then
                        If Len(colTagValues.Item(CONST_TagSuffixAttributeName)) > 0 Then
                            .Attributes.Item(CONST_TagSuffixAttributeName).Value = colTagValues.Item(CONST_TagSuffixAttributeName)
                        End If
                    End If
                End If

                If Not IsDBNull(varValue) Then
                    'trim off any spaces in the input value
                    varValue = Trim(varValue)
                End If
                .Commit()
            End With
        End If

        Err.Clear()

    End Function

    Private Function GetMatchingLoopTag(Datasource As LMADataSource, Item As Llama.LMInstrument, sMeasuredVariable As String, sTagSequenceNo As String) As LMInstrLoop
        Dim sCatalogExplorerRootPath As String
        Dim i As Long
        Dim vData As Object
        Dim DomObj As XmlDocument
        Dim objTopNodeList As XmlNodeList
        Dim objEntityEle As XmlElement
        Dim objXmlAtt As XmlNode
        Dim blnFound As Boolean
        Dim strItemTagFormat As String
        Dim sLoopItemTag As String
        Dim sNextLoopItemTag As String
        Dim objLMInstrLoop As LMInstrLoop
        Dim objTmpCriterion As LMACriterion
        Dim objFilter As LMAFilter
        Dim objLMInstrLoops As LMInstrLoops
        Dim LoopFormats As String() = Nothing

        DomObj = New XmlDocument
        objFilter = New LMAFilter
        objLMInstrLoops = New LMInstrLoops

        objFilter.ItemType = "InstrLoop"
        objFilter.Conjunctive = True

        On Error GoTo Other

        GetMatchingLoopTag = Nothing

        sCatalogExplorerRootPath = GetOptionManagerValue(Datasource, CONST_NameAttributeName, CONST_CatalogExplorerrootpath)
        DomObj.Load(Left(sCatalogExplorerRootPath, InStr(1, sCatalogExplorerRootPath, "Symbols") - 2) & "\ItemTagFormat.xml")

        objTopNodeList = DomObj.GetElementsByTagName(Const_Entity)

        For Each objEntityEle In objTopNodeList
            For Each objXmlAtt In objEntityEle.Attributes
                If objXmlAtt.Name = Const_ItemType Then
                    If objXmlAtt.Value = CONST_InstrLoopItemName Then
                        blnFound = True
                        strItemTagFormat = objEntityEle.GetElementsByTagName("Attribute").Item(0).Attributes(0).InnerText
                        ReDim LoopFormats(0)
                        LoopFormats = Split(strItemTagFormat, "(")
                        Exit For
                    End If
                End If
            Next
            If blnFound Then Exit For
        Next

        If blnFound Then
            Dim tmpSourceAttrName As String
            For i = 0 To UBound(LoopFormats)
                If LoopFormats(i) <> vbNullString Then
                    tmpSourceAttrName = Left(LoopFormats(i), InStr(1, LoopFormats(i), ")") - 1)
                    objTmpCriterion = New LMACriterion
                    objTmpCriterion.Conjunctive = CONST_MatchAll
                    If InStr(1, tmpSourceAttrName, "...") Then
                        objTmpCriterion.SourceAttributeName = "SP_PlantGroupID"
                        objTmpCriterion.Operator = "="
                        objTmpCriterion.ValueAttribute = Item.PlantGroupID
                        objTmpCriterion.Bind = True
                    Else
                        objTmpCriterion.SourceAttributeName = tmpSourceAttrName
                        objTmpCriterion.Operator = "="
                        If tmpSourceAttrName = "LoopFunction" Then tmpSourceAttrName = CONST_MeasuredVariableCodeAttributeName
                        objTmpCriterion.ValueAttribute = sMeasuredVariable
                        objTmpCriterion.Bind = True
                    End If
                    If objFilter.Criteria.Add(objTmpCriterion) = False Then
                        LogAndRaiseError("::GetMatchingLoopTag()" & " Could not add criteria.")
                    End If
                End If
            Next i
        Else

Other:
            On Error GoTo ErrHandler
            If Len(Err.Description) > 0 Then
                LogError("::GetMatchingLoopTag()" & Err.Description)
                'clear filter
                objFilter = New LMAFilter
            End If

            objTmpCriterion = New LMACriterion
            objTmpCriterion.SourceAttributeName = CONST_TagSequenceNoAttributeName
            objTmpCriterion.Operator = "="
            objTmpCriterion.ValueAttribute = sTagSequenceNo
            objTmpCriterion.Bind = True
            If objFilter.Criteria.Add(objTmpCriterion) = False Then
                LogAndRaiseError("::GetMatchingLoopTag()" & " Could not add criteria.")
            End If

            If Not sMeasuredVariable = "" Then
                objTmpCriterion = New LMACriterion
                objTmpCriterion.Conjunctive = CONST_MatchAll 'Match all
                objTmpCriterion.SourceAttributeName = CONST_LoopFunctionAttributeName
                objTmpCriterion.Operator = "="
                objTmpCriterion.ValueAttribute = sMeasuredVariable
                If objFilter.Criteria.Add(objTmpCriterion) = False Then
                    LogAndRaiseError("::GetMatchingLoopTag()" & " Could not add criteria1.")
                End If
            End If
        End If

        'Set the filter
        objLMInstrLoops.Collect(Datasource, Filter:=objFilter)

        If objLMInstrLoops.Count > 0 Then
            sLoopItemTag = objLMInstrLoops.Nth(1).Attributes("ItemTag").Value
            objLMInstrLoop = objLMInstrLoops.Nth(1)

            For i = 2 To objLMInstrLoops.Count
                sNextLoopItemTag = objLMInstrLoops.Nth(i).Attributes("ItemTag").Value
                If sLoopItemTag > sNextLoopItemTag Then
                    sLoopItemTag = sNextLoopItemTag
                    objLMInstrLoop = objLMInstrLoops.Nth(i)
                End If
            Next

            If CheckItemStatusAndStockPileStatus(Datasource, Item, objLMInstrLoop.Id) Then
                If CheckIsSharedAndSiteId(Datasource, Item, objLMInstrLoop.Id) Then
                    GetMatchingLoopTag = objLMInstrLoop
                End If
            End If
        End If

        'Cleanup
        objFilter = Nothing
        objTmpCriterion = Nothing
        objLMInstrLoops = Nothing
        '
        Exit Function
ErrHandler:
        'Cleanup
        objFilter = Nothing
        objTmpCriterion = Nothing
        objLMInstrLoops = Nothing
        LogAndRaiseError("::GetMatchingLoopTag()")
    End Function

    Private Function CheckIsSharedAndSiteId(Datasource As LMADataSource, Item As Llama.LMInstrument, strSPID As String) As Boolean
        Dim objCriterion As LMACriterion
        Dim objFilter As LMAFilter
        Dim objPlantItemGroup As LMPlantItemGroup
        Dim objPlantItemGroups As LMPlantItemGroups

        On Error GoTo ErrHandler
        CheckIsSharedAndSiteId = False

        objFilter = New LMAFilter
        objPlantItemGroups = New LMPlantItemGroups

        objCriterion = New LMACriterion
        objCriterion.SourceAttributeName = CONST_SP_IDAttributeName
        objCriterion.Operator = "="
        objCriterion.ValueAttribute = strSPID
        objCriterion.Bind = True
        If objFilter.Criteria.Add(objCriterion) = False Then
            LogAndRaiseError("::CheckIsSharedAndSiteId()" & " Could not add criteria.")
        End If

        'Set the filter
        objPlantItemGroups.Collect(Datasource, Filter:=objFilter)

        If objPlantItemGroups.Count > 0 Then
            objPlantItemGroup = objPlantItemGroups.Nth(1)

            'check if NonWorkShare
            If Datasource.IPile.WorkSharePlantType <> 3 Then
                If CBool(objPlantItemGroup.Attributes("IsShared").value) Then
                    CheckIsSharedAndSiteId = True
                ElseIf CStr(objPlantItemGroup.Attributes("SP_WSSiteID").value) = CStr(Datasource.IPile.ActiveWorkShareSiteID) Then
                    CheckIsSharedAndSiteId = True
                End If
            Else
                CheckIsSharedAndSiteId = True
            End If
        End If

        'Cleanup
        objFilter = Nothing
        objCriterion = Nothing
        objPlantItemGroup = Nothing
        objPlantItemGroups = Nothing

        Exit Function

ErrHandler:
        'Cleanup
        objFilter = Nothing
        objCriterion = Nothing
        objPlantItemGroup = Nothing
        objPlantItemGroups = Nothing
        LogAndRaiseError("::CheckIsSharedAndSiteId()")
    End Function

    Private Function CheckItemStatusAndStockPileStatus(Datasource As LMADataSource, Item As Llama.LMInstrument, strSPID As String) As Boolean
        Dim i As Long
        Dim objCriterion As LMACriterion
        Dim objFilter As LMAFilter
        Dim objRepresentation As LMRepresentation
        Dim objRepresentations As LMRepresentations

        On Error GoTo ErrHandler

        CheckItemStatusAndStockPileStatus = False

        objFilter = New LMAFilter
        objRepresentations = New LMRepresentations

        objCriterion = New LMACriterion
        objCriterion.SourceAttributeName = CONST_ModelItemIDAttribute
        objCriterion.Operator = "="
        objCriterion.ValueAttribute = strSPID
        objCriterion.Bind = True
        If objFilter.Criteria.Add(objCriterion) = False Then
            LogAndRaiseError("::CheckItemStatusAndStockPileStatus()" & " Could not add criteria.")
        End If

        objCriterion = New LMACriterion
        objCriterion.Conjunctive = CONST_MatchAll 'Match all
        objCriterion.SourceAttributeName = CONST_ItemStatusAttribute
        objCriterion.Operator = "<>"
        objCriterion.ValueAttribute = "4"
        If objFilter.Criteria.Add(objCriterion) = False Then
            LogAndRaiseError("::CheckItemStatusAndStockPileStatus()" & " Could not add criteria1.")
        End If

        'Set the filter
        objRepresentations.Collect(Datasource, Filter:=objFilter)

        If objRepresentations.Count > 0 Then
            objRepresentation = objRepresentations.Nth(1)

            'in case we will have more then one representation for the instrument in the future.
            For i = 1 To Item.Representations.Count
                If Item.Representations.Nth(i).DrawingID = objRepresentation.DrawingID Or objRepresentation.DrawingID = 0 Then
                    CheckItemStatusAndStockPileStatus = True
                    Exit For
                End If
            Next
        End If

        'Cleanup
        objFilter = Nothing
        objCriterion = Nothing
        objRepresentation = Nothing
        objRepresentations = Nothing

        Exit Function

ErrHandler:
        'Cleanup
        objFilter = Nothing
        objCriterion = Nothing
        objRepresentation = Nothing
        objRepresentations = Nothing
        LogAndRaiseError("::CheckItemStatusAndStockPileStatus()")
    End Function

    'Function to get value from Options Manager
    Private Function GetOptionManagerValue(ByVal objDataSource As LMADataSource, ByVal SourceAttributeName As String, ByVal ValueAttribute As String) As String
        On Error GoTo ErrorHandler

        Dim objOptionSettings As New LMOptionSettings
        Dim objFilter As LMAFilter

        objOptionSettings = New LMOptionSettings

        'Forming a filter
        objFilter = GetFilter(CONST_OptionSettingAttributeName, SourceAttributeName, ValueAttribute)

        With objDataSource
            .QueryCache = False

            With objOptionSettings
                .Collect(objDataSource, Filter:=objFilter)
                GetOptionManagerValue = Trim$(.Nth(1).Value)
            End With

            .QueryCache = True
        End With

Cleanup:
        objOptionSettings = Nothing
        objFilter = Nothing

        Exit Function

ErrorHandler:
        LogError("ItemTag - ItemTagFunc::GetOptionManagerValue --> " & Err.Description)
        Resume Cleanup

    End Function

    Private Function UpdatePipeRunTag(ByRef Datasource As Llama.LMADataSource, ByRef Item As Llama.LMPipeRun, ByRef varValue As Object, ByRef strPropertyName As String) As Boolean

        Dim colTagValues As Collection
        Dim objPlantGroups As Llama.LMPlantGroups
        Dim objPlantGroup As Llama.LMPlantGroup
        Dim objUnit As Llama.LMUnit
        Dim objFilter As Llama.LMAFilter
        Dim objPipeRun As Llama.LMPipeRun

        Dim strItemTag As String
        Dim strLocTagSeqNo As String
        Dim strDupMessage As String
        Dim strUnitName As String
        Dim intResponse As Short
        Dim blnUnique As Boolean
        Dim bUniqueStatus As Boolean

        On Error GoTo ErrorHandler

        strUnitName = ""
        strDupMessage = ""

        Item.Attributes.BuildAttributesOnDemand = True

        'Get the UnitCode Value
        objPipeRun = Datasource.GetPipeRun((Item.Id))
        objPipeRun.Attributes.BuildAttributesOnDemand = True

        On Error Resume Next
        objPlantGroup = objPipeRun.PlantGroupObject
        On Error GoTo ErrorHandler
        If Not objPlantGroup Is Nothing Then
            objPlantGroup.Attributes.BuildAttributesOnDemand = True

            If objPlantGroup.PlantGroupTypeIndex = 65 Then
                objUnit = Datasource.GetUnit((objPlantGroup.Id))
                objUnit.Attributes.BuildAttributesOnDemand = True
                'Unit code may not be set if the Units are created by retrieving from EF
                If Not IsDBNull(objUnit.UnitCode) Then
                    strUnitName = objUnit.UnitCode
                End If
            End If
        End If

        'Use the collections unique index feature to avoid a bunch of if statements to
        'determine what property was passed in
        colTagValues = New Collection

        With colTagValues

            .Add(Trim(VariantToString(varValue)), strPropertyName)

            Select Case strPropertyName

                Case CONST_TagSuffixAttributeName
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)
                    .Add(Trim(VariantToString((Item.OperFluidCode))), CONST_OperFluidCodeAttributeName)

                Case CONST_TagSequenceNoAttributeName
                    .Add(Trim(VariantToString((Item.TagSuffix))), CONST_TagSuffixAttributeName)
                    .Add(Trim(VariantToString((Item.OperFluidCode))), CONST_OperFluidCodeAttributeName)

                Case CONST_OperFluidCodeAttributeName
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)
                    .Add(Trim(VariantToString((Item.TagSuffix))), CONST_TagSuffixAttributeName)

                Case Else
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)
                    .Add(Trim(VariantToString((Item.TagSuffix))), CONST_TagSuffixAttributeName)
                    .Add(Trim(VariantToString((Item.OperFluidCode))), CONST_OperFluidCodeAttributeName)

            End Select

        End With

        strLocTagSeqNo = colTagValues.Item(CONST_TagSequenceNoAttributeName)

        If Len(strLocTagSeqNo) = 0 And strPropertyName <> CONST_TagSequenceNoAttributeName Then
            'Get the next tagseqno from Options Manager
            strLocTagSeqNo = GetNextAvailTagSeqNo(Datasource, CONST_NameAttributeName, CONST_PipeRunNextSeqNoAttributeName)
        End If

        strItemTag = BuildPipeRunTag(strUnitName, strLocTagSeqNo, colTagValues.Item(CONST_TagSuffixAttributeName), colTagValues.Item(CONST_OperFluidCodeAttributeName))

        UpdatePipeRunTag = True

        If Len(strLocTagSeqNo) > 0 Then
            objFilter = Nothing
            strDupMessage = My.Resources.str5000
            If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuiltAndProjs Then
                strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Pipe Run " & My.Resources.str5011
                If Item.IsBulkItemIndex = CONST_TrueIndex Then
                    objFilter = GetFilter(CONST_PipeRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_PipeRunPlantItemIndex), CONST_IsBulkItemAttributeName, CStr(CONST_FalseIndex))
                Else
                    objFilter = GetFilter(CONST_PipeRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_PipeRunPlantItemIndex))
                End If
                bUniqueStatus = CheckForGlobalUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex), strItemTag, "")
            Else
                If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuilt Then
                    'strDupMessage = "Duplicate Tag " & strItemTag & " found in active project."
                    strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Pipe Run " & My.Resources.str5010
                Else
                    'strDupMessage = "Duplicate Tag " & strItemTag & " found in active plant."
                    strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Pipe Run " & My.Resources.str5007
                End If

                'Check Active Plant first
                objFilter = GetFilter(CONST_PipeRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_ItemStatus, Const_ItemStatusValue)
                bUniqueStatus = CheckForUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex))

                If bUniqueStatus = True Then
                    'strDupMessage = "Duplicate Tag " & strItemTag & " found in AsBuilt plant."
                    strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Pipe Run " & My.Resources.str5008
                    If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuilt Then
                        objFilter = Nothing
                        If Item.IsBulkItemIndex = CONST_TrueIndex Then
                            objFilter = GetFilter(CONST_PipeRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_PipeRunPlantItemIndex), CONST_IsBulkItemAttributeName, CStr(CONST_FalseIndex))
                        Else
                            objFilter = GetFilter(CONST_PipeRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_PipeRunPlantItemIndex))
                        End If
                        bUniqueStatus = CheckForAsBuiltUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex), strItemTag, "")
                    End If
                End If
            End If

            If bUniqueStatus = False Then
                If m_isUIEnabled Then
                    intResponse = MsgBox(strDupMessage & " " & My.Resources.str5009, MsgBoxStyle.YesNo + MsgBoxStyle.SystemModal + MsgBoxStyle.Exclamation, My.Resources.str5001)
                Else
                    intResponse = MsgBoxResult.No
                End If

                If intResponse = MsgBoxResult.No Then

                    UpdatePipeRunTag = True

                    With Item
                        If Len(strLocTagSeqNo) > 0 Then
                            .Attributes.Item(CONST_TagSequenceNoAttributeName).Value = strLocTagSeqNo
                        End If

                        If Len(colTagValues.Item(CONST_TagSuffixAttributeName)) > 0 Then
                            .Attributes.Item(CONST_TagSuffixAttributeName).Value = colTagValues.Item(CONST_TagSuffixAttributeName)
                        End If
                        If Len(strItemTag) > 0 Then
                            .ItemTag = strItemTag
                        End If
                        .Commit()
                    End With

                    Exit Function

                ElseIf intResponse = MsgBoxResult.Yes Then

                    blnUnique = False

                    Do While blnUnique = False

                        'Get the next tagseqno from Options Manager
                        strLocTagSeqNo = GetNextAvailTagSeqNo(Datasource, CONST_NameAttributeName, CONST_PipeRunNextSeqNoAttributeName)

                        strItemTag = BuildPipeRunTag(strUnitName, strLocTagSeqNo, colTagValues.Item(CONST_TagSuffixAttributeName), colTagValues.Item(CONST_OperFluidCodeAttributeName))

                        objFilter = Nothing
                        If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuiltAndProjs Then
                            strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Pipe Run " & My.Resources.str5011
                            If Item.IsBulkItemIndex = CONST_TrueIndex Then
                                'if the piperun is bulk its okay to have a duplicate on another bulk
                                'but not on a not-bulk.
                                objFilter = GetFilter(CONST_PipeRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_PipeRunPlantItemIndex), CONST_IsBulkItemAttributeName, CStr(CONST_FalseIndex))
                            Else
                                objFilter = GetFilter(CONST_PipeRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_PipeRunPlantItemIndex))
                            End If
                            blnUnique = CheckForGlobalUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex), strItemTag, "")
                        Else
                            objFilter = GetFilter(CONST_PipeRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_ItemStatus, Const_ItemStatusValue)
                            blnUnique = CheckForUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex))

                            If blnUnique = True Then
                                'Its unique in the project - check in the Asbuilt
                                If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuilt Then
                                    objFilter = Nothing
                                    If Item.IsBulkItemIndex = CONST_TrueIndex Then
                                        objFilter = GetFilter(CONST_PipeRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_PipeRunPlantItemIndex), CONST_IsBulkItemAttributeName, CStr(CONST_FalseIndex))
                                    Else
                                        objFilter = GetFilter(CONST_PipeRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_PipeRunPlantItemIndex))
                                    End If
                                    blnUnique = CheckForAsBuiltUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex), strItemTag, "")
                                End If
                            End If
                        End If
                    Loop

                Else
                    UpdatePipeRunTag = False
                End If

            End If

        End If

        'Update the item values
        If UpdatePipeRunTag Then

            With Item

                If Len(strLocTagSeqNo) > 0 Then
                    .Attributes.Item(CONST_TagSequenceNoAttributeName).Value = strLocTagSeqNo
                Else
                    .Attributes.Item(CONST_TagSequenceNoAttributeName).Value = System.DBNull.Value
                End If

                If strLocTagSeqNo = "" Then
                    If Len(colTagValues.Item(CONST_OperFluidCodeAttributeName)) = 0 Then
                        .ItemTag = System.DBNull.Value
                    Else
                        If Len(colTagValues.Item(CONST_TagSuffixAttributeName)) > 0 Then
                            .Attributes.Item(CONST_TagSuffixAttributeName).Value = colTagValues.Item(CONST_TagSuffixAttributeName)
                        End If
                        If Len(strItemTag) > 0 Then
                            .ItemTag = strItemTag
                        End If
                    End If

                    .ItemTag = System.DBNull.Value
                Else
                    If Len(colTagValues.Item(CONST_TagSuffixAttributeName)) > 0 Then
                        .Attributes.Item(CONST_TagSuffixAttributeName).Value = colTagValues.Item(CONST_TagSuffixAttributeName)
                    End If
                    If Len(strItemTag) > 0 Then
                        .ItemTag = strItemTag
                    End If
                End If

                If Not IsDBNull(varValue) Then
                    'trim off any spaces in the input value
                    varValue = Trim(varValue)
                End If

                .Commit()

            End With

        End If

        Exit Function

ErrorHandler:
        LogError("ItemTag - ItemTagFunc::UpdatePipeRunTag --> " & Err.Description)
        objPipeRun = Nothing
        objPlantGroups = Nothing
        objUnit = Nothing
        colTagValues = Nothing
        objFilter = Nothing

    End Function

    Private Function UpdateEquipTag(ByRef Datasource As Llama.LMADataSource, ByRef Item As Llama.LMEquipment, ByRef varValue As Object, ByRef strPropertyName As String) As Boolean

        Dim colTagValues As Collection
        Dim objFilter As Llama.LMAFilter
        Dim strDupMessage As String
        Dim strItemTag As String
        Dim strLocTagSeqNo As String
        Dim intResponse As Short
        Dim blnUnique As Boolean
        Dim bUniqueStatus As Boolean

        On Error GoTo ErrorHandler

        Item.Attributes.BuildAttributesOnDemand = True
        strDupMessage = ""

        'Use the collections unique index feature to avoid a bunch of if statements
        'to determine what property was passed in
        colTagValues = New Collection

        With colTagValues
            'Trim off any white spaces in the properties that make up the tag.
            .Add(Trim(VariantToString(varValue)), strPropertyName)

            Select Case strPropertyName

                Case CONST_TagSuffixAttributeName
                    .Add(Trim(VariantToString((Item.TagPrefix))), CONST_TagPrefixAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)

                Case CONST_TagPrefixAttributeName
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)
                    .Add(Trim(VariantToString((Item.TagSuffix))), CONST_TagSuffixAttributeName)

                Case CONST_TagSequenceNoAttributeName
                    .Add(Trim(VariantToString((Item.TagPrefix))), CONST_TagPrefixAttributeName)
                    .Add(Trim(VariantToString((Item.TagSuffix))), CONST_TagSuffixAttributeName)

                Case CONST_ItemTagAttributeName
                    .Add(Trim(VariantToString((Item.TagPrefix))), CONST_TagPrefixAttributeName)
                    .Add(Trim(VariantToString((Item.TagSuffix))), CONST_TagSuffixAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)
            End Select

        End With

        strLocTagSeqNo = colTagValues.Item(CONST_TagSequenceNoAttributeName)

        If Len(strLocTagSeqNo) < 1 And strPropertyName <> CONST_TagSequenceNoAttributeName And (Trim(VariantToString(varValue)) <> "" Or strPropertyName = CONST_ItemTagAttributeName) Then

            'Get the next tagseqno from Options Manager
            strLocTagSeqNo = GetNextAvailTagSeqNo(Datasource, CONST_NameAttributeName, CONST_EquipNextSeqNoAttributeName)

            strItemTag = BuildEquipmentTag(colTagValues.Item(CONST_TagPrefixAttributeName), strLocTagSeqNo, colTagValues.Item(CONST_TagSuffixAttributeName))

            bUniqueStatus = CheckForEquipUnique(Datasource, Item, strItemTag, strDupMessage)

            'Check for uniqueness of the ItemTag
            If bUniqueStatus = False Then
                blnUnique = False

                Do While blnUnique = False

                    'Get the next tagseqno from Options Manager
                    strLocTagSeqNo = GetNextAvailTagSeqNo(Datasource, CONST_NameAttributeName, CONST_EquipNextSeqNoAttributeName)

                    strItemTag = BuildEquipmentTag(colTagValues.Item(CONST_TagPrefixAttributeName), strLocTagSeqNo, colTagValues.Item(CONST_TagSuffixAttributeName))

                    blnUnique = CheckForEquipUnique(Datasource, Item, strItemTag, strDupMessage)
                Loop
            End If
        End If

        strItemTag = BuildEquipmentTag(colTagValues.Item(CONST_TagPrefixAttributeName), strLocTagSeqNo, colTagValues.Item(CONST_TagSuffixAttributeName))

        UpdateEquipTag = True

        If Len(strLocTagSeqNo) > 0 Then
            'Check for uniqueness
            bUniqueStatus = CheckForEquipUnique(Datasource, Item, strItemTag, strDupMessage)

            If bUniqueStatus = False Then

                If m_isUIEnabled Then
                    intResponse = MsgBox(strDupMessage & " " & My.Resources.str5009, MsgBoxStyle.YesNo + MsgBoxStyle.SystemModal + MsgBoxStyle.Exclamation, My.Resources.str5001)
                Else
                    intResponse = MsgBoxResult.Yes
                End If

                If intResponse = MsgBoxResult.Yes Then
                    blnUnique = False

                    Do While blnUnique = False

                        'Get the next Available tag sequence number from Options Manager
                        strLocTagSeqNo = GetNextAvailTagSeqNo(Datasource, CONST_NameAttributeName, CONST_EquipNextSeqNoAttributeName)

                        strItemTag = BuildEquipmentTag(colTagValues.Item(CONST_TagPrefixAttributeName), strLocTagSeqNo, colTagValues.Item(CONST_TagSuffixAttributeName))

                        blnUnique = CheckForEquipUnique(Datasource, Item, strItemTag, strDupMessage)
                    Loop
                Else
                    varValue = Item.Attributes.Item(strPropertyName).Value
                    UpdateEquipTag = False
                End If

            End If

        End If

        'Update the values
        If UpdateEquipTag Then

            With Item
                If Len(strLocTagSeqNo) > 0 Then
                    .Attributes.Item(CONST_TagSequenceNoAttributeName).Value = strLocTagSeqNo
                Else
                    .Attributes.Item(CONST_TagSequenceNoAttributeName).Value = System.DBNull.Value
                End If

                If strLocTagSeqNo = "" Then
                    .ItemTag = System.DBNull.Value
                Else
                    If Len(colTagValues.Item(CONST_TagPrefixAttributeName)) > 0 Then
                        .TagPrefix = colTagValues.Item(CONST_TagPrefixAttributeName)
                    End If
                    If Len(colTagValues.Item(CONST_TagSuffixAttributeName)) > 0 Then
                        .TagSuffix = colTagValues.Item(CONST_TagSuffixAttributeName)
                    End If

                    If Len(strItemTag) > 0 Then
                        .ItemTag = strItemTag
                    End If
                End If

                If Not IsDBNull(varValue) Then
                    'trim off any spaces in the input value
                    varValue = Trim(varValue)
                End If
                .Commit()

            End With

        End If

        Exit Function

ErrorHandler:
        LogError("ItemTag - ItemTagFunc::UpdateEquipTag --> " & Err.Description)
        colTagValues = Nothing
        objFilter = Nothing

    End Function

    Private Function CheckForEquipUnique(ByRef Datasource As Llama.LMADataSource, ByRef Item As Llama.LMEquipment, ByRef strItemTag As String, ByRef strMessage As String) As Boolean
        Dim bUniqueStatus As Boolean
        Dim objFilter As Llama.LMAFilter
        Dim strDupMessage As String

        On Error GoTo ErrorHandler
        CheckForEquipUnique = False
        bUniqueStatus = False
        strMessage = ""
        strDupMessage = ""

        If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuiltAndProjs Then
            strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Equipment " & My.Resources.str5011 'AsBuiltPlant or its projects
            If Len(Item.PartOfPlantItemID.ToString()) > 0 Then
                If Item.IsBulkItemIndex = CONST_TrueIndex Then
                    objFilter = GetFilter(CONST_EquipmentItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PartofPlantItem_SP_IDAttributeName, Item.PartOfPlantItemID, CONST_PlantItemTypeAttributeName, CStr(CONST_EquipmentPlantItemIndex), CONST_IsBulkItemAttributeName, CStr(CONST_FalseIndex))
                Else
                    objFilter = GetFilter(CONST_EquipmentItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PartofPlantItem_SP_IDAttributeName, Item.PartOfPlantItemID, CONST_PlantItemTypeAttributeName, CStr(CONST_EquipmentPlantItemIndex))
                End If
            Else
                If Item.IsBulkItemIndex = CONST_TrueIndex Then
                    objFilter = GetFilter(CONST_EquipmentItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_EquipmentPlantItemIndex), CONST_IsBulkItemAttributeName, CStr(CONST_FalseIndex))
                Else
                    objFilter = GetFilter(CONST_EquipmentItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_EquipmentPlantItemIndex))
                End If
            End If
            bUniqueStatus = CheckForGlobalUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex), strItemTag, "")
        Else
            'Check the local plant or project
            If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuilt Then
                strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Equipment " & My.Resources.str5010 'Active project
            Else
                strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Equipment " & My.Resources.str5007 'Active plant
            End If
            If Len(Item.PartOfPlantItemID.ToString()) > 0 Then
                objFilter = GetFilter(CONST_EquipmentItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PartofPlantItem_SP_IDAttributeName, Item.PartOfPlantItemID, CONST_ItemStatus, Const_ItemStatusValue)
            Else
                objFilter = GetFilter(CONST_EquipmentItemName, CONST_ItemTagAttributeName, strItemTag, CONST_ItemStatus, Const_ItemStatusValue)
            End If
            bUniqueStatus = CheckForUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex))

            'Do we need to check the AsBuilt too?
            If bUniqueStatus = True Then
                If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuilt Then
                    strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Equipment " & My.Resources.str5008 'found in AsBuilt Plant
                    If Len(Item.PartOfPlantItemID.ToString()) > 0 Then
                        If Item.IsBulkItemIndex = CONST_TrueIndex Then
                            objFilter = GetFilter(CONST_EquipmentItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PartofPlantItem_SP_IDAttributeName, Item.PartOfPlantItemID, CONST_PlantItemTypeAttributeName, CStr(CONST_EquipmentPlantItemIndex), CONST_IsBulkItemAttributeName, CStr(CONST_FalseIndex))
                        Else
                            objFilter = GetFilter(CONST_EquipmentItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PartofPlantItem_SP_IDAttributeName, Item.PartOfPlantItemID, CONST_PlantItemTypeAttributeName, CStr(CONST_EquipmentPlantItemIndex))
                        End If
                    Else
                        If Item.IsBulkItemIndex = CONST_TrueIndex Then
                            objFilter = GetFilter(CONST_EquipmentItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_EquipmentPlantItemIndex), CONST_IsBulkItemAttributeName, CStr(CONST_FalseIndex))
                        Else
                            objFilter = GetFilter(CONST_EquipmentItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_EquipmentPlantItemIndex))
                        End If
                    End If
                    bUniqueStatus = CheckForAsBuiltUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex), strItemTag, "")
                End If
            End If
        End If

        CheckForEquipUnique = bUniqueStatus
        If CheckForEquipUnique = False Then
            strMessage = strDupMessage
        End If
        Exit Function

ErrorHandler:
        LogError("ItemTag - ItemTagFunc::CheckForEquipUnique --> " & Err.Description)
        objFilter = Nothing

    End Function

    Private Function CheckForInstrUnique(ByRef Datasource As Llama.LMADataSource, ByRef Item As Llama.LMInstrument, ByRef strItemTag As String, ByRef strMessage As String) As Boolean
        Dim bUniqueStatus As Boolean
        Dim objFilter As Llama.LMAFilter
        Dim strDupMessage As String

        On Error GoTo ErrorHandler
        CheckForInstrUnique = False
        bUniqueStatus = False
        strMessage = ""
        strDupMessage = ""

        'Forming a filter
        If (m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuiltAndProjs) Then
            strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Instrument " & My.Resources.str5011 'AsBuilt plant or its projects
            'Get a filter to use for the V_GlobalTagList view  (this is not a table)
            If Item.IsBulkItemIndex = CONST_TrueIndex Then
                objFilter = GetFilter(CONST_InstrumentItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_InstrumentPlantItemIndex1), CONST_IsBulkItemAttributeName, CStr(CONST_FalseIndex))
            Else
                objFilter = GetFilter(CONST_InstrumentItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_InstrumentPlantItemIndex1))
            End If

            'CheckForGlobalUnique
            bUniqueStatus = CheckForGlobalUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex), strItemTag, "")
        Else
            If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuilt Then
                strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Instrument " & My.Resources.str5010 'Active Project
            Else
                strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Instrument " & My.Resources.str5007 'Active Plant
            End If
            objFilter = GetFilter(CONST_InstrumentItemName, CONST_ItemTagAttributeName, strItemTag, CONST_ItemStatus, Const_ItemStatusValue)

            'Checking if the itemtag is unique in this project
            bUniqueStatus = CheckForUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex))

            If bUniqueStatus = True Then
                If (m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuilt) Then
                    strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Instrument " & My.Resources.str5008 'AsBuilt Plant
                    objFilter = Nothing
                    If Item.IsBulkItemIndex = CONST_TrueIndex Then
                        objFilter = GetFilter(CONST_InstrumentItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_InstrumentPlantItemIndex1), CONST_IsBulkItemAttributeName, CStr(CONST_FalseIndex))
                    Else
                        objFilter = GetFilter(CONST_InstrumentItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_InstrumentPlantItemIndex1))
                    End If
                    'CheckForAsBuiltUnique
                    bUniqueStatus = CheckForAsBuiltUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex), strItemTag, "")
                End If
            End If
        End If

        CheckForInstrUnique = bUniqueStatus
        If CheckForInstrUnique = False Then
            strMessage = strDupMessage
        End If
        Exit Function

ErrorHandler:
        LogError("ItemTag - ItemTagFunc::CheckForInstrUnique --> " & Err.Description)
        objFilter = Nothing

    End Function

    Private Function UpdateNozzleTag(ByRef Datasource As Llama.LMADataSource, ByRef Item As Llama.LMNozzle, ByRef varValue As Object, ByRef strPropertyName As String) As Boolean

        Dim objLMPlantItem As Llama.LMPlantItem
        Dim objChildPlantItem As Llama.LMPlantItem
        Dim objLMEquipment As Llama.LMEquipment
        Dim objEquipComp As Llama.LMEquipComponents
        Dim objLMNozzle As Llama.LMNozzle
        Dim vPartOfID As Object
        Dim strEquipID As String
        Dim ItemTag_Renamed As String
        Dim locTagSeqNo As String
        Dim nBulkItemIndex As Integer
        Dim Response As Short
        Dim colTagValues As Collection
        Dim strDupMessage As String
        Dim strEquipTag As String
        Dim strLocTagSeqNo As String
        Dim strItemTag As String
        Dim intResponse As Short
        Dim DoUntilUni As Boolean
        Dim objlmaitems As Llama.LMAItems
        Dim objFilter As Llama.LMAFilter
        Dim strLocTagPrefix As String
        Dim strLocTagSuffix As String
        Dim bAttachToRoom As Boolean

        On Error GoTo errHandler

        If IsRoomAttached(Item, Datasource) Then
            UpdateNozzleTag = UpdateNozzleTagForRoom(Datasource, Item, varValue, strPropertyName)
            Exit Function
        End If

        UpdateNozzleTag = True

        'In following cases do not increment TagSeqNo
        '1.Placement
        '2.Drag&Drop when nozzle don't have item tag
        If strPropertyName = CONST_ItemTagAttributeName And Len("" & varValue) = 0 Then Exit Function

        colTagValues = New Collection
        strDupMessage = ""
        strEquipTag = ""

        Item.Attributes.BuildAttributesOnDemand = True

        With colTagValues
            .Add(Trim(VariantToString(varValue)), strPropertyName)

            Select Case strPropertyName

                Case CONST_TagSuffixAttributeName
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagPrefixAttributeName).Value))), CONST_TagPrefixAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)

                Case CONST_TagPrefixAttributeName
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSuffixAttributeName).Value))), CONST_TagSuffixAttributeName)

                Case CONST_TagSequenceNoAttributeName
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagPrefixAttributeName).Value))), CONST_TagPrefixAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSuffixAttributeName).Value))), CONST_TagSuffixAttributeName)

                Case CONST_ItemTagAttributeName
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagPrefixAttributeName).Value))), CONST_TagPrefixAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSuffixAttributeName).Value))), CONST_TagSuffixAttributeName)

                Case Else
            End Select
        End With

        strLocTagSeqNo = colTagValues.Item(CONST_TagSequenceNoAttributeName)
        strItemTag = BuildNozzleTag(colTagValues.Item(CONST_TagPrefixAttributeName), strLocTagSeqNo, colTagValues.Item(CONST_TagSuffixAttributeName))

        If strPropertyName = CONST_TagSequenceNoAttributeName And strLocTagSeqNo = "" Then GoTo ExitDueTOTagSeq
        If strLocTagSeqNo = "" Then
            If colTagValues.Item(CONST_TagSuffixAttributeName) <> "" Then
                DoUntilUni = True
            Else
                If colTagValues.Item(CONST_TagPrefixAttributeName) <> "" Then
                    DoUntilUni = True
                End If
            End If
        End If

        strLocTagPrefix = colTagValues.Item(CONST_TagPrefixAttributeName)
        If strPropertyName = CONST_TagPrefixAttributeName And strLocTagPrefix <> "" Then
            If strLocTagSeqNo = "" Then
                strLocTagSeqNo = CStr(1)
            End If
        End If
        strLocTagSuffix = colTagValues.Item(CONST_TagSuffixAttributeName)
        If strPropertyName = CONST_TagSuffixAttributeName And strLocTagSuffix <> "" Then
            If strLocTagSeqNo = "" Then
                strLocTagSeqNo = CStr(1)
            End If
        End If

        strItemTag = BuildNozzleTag(colTagValues.Item(CONST_TagPrefixAttributeName), strLocTagSeqNo, colTagValues.Item(CONST_TagSuffixAttributeName))

        nBulkItemIndex = Item.IsBulkItemIndex

        'get the EquipmentID from the Nozzle. Here Nozzle can be on Equipment Component.
        'so navigate upto the Equipment.
        If IsDBNull(Item.EquipmentID) Then

            objLMPlantItem = Datasource.GetPlantItem((Item.Id))
            objLMPlantItem.Attributes.BuildAttributesOnDemand = True

            vPartOfID = objLMPlantItem.PartOfPlantItemID

            While Not IsDBNull(vPartOfID)
                objLMPlantItem = objLMPlantItem.PartOfPlantItemObject
                objLMPlantItem.Attributes.BuildAttributesOnDemand = True

                vPartOfID = objLMPlantItem.PartOfPlantItemID
            End While
            If objLMPlantItem.PlantItemTypeIndex = 21 Then
                strEquipID = CStr(objLMPlantItem.Attributes(CONST_SP_EquipmentIDAttributeName).Value)
            Else
                strEquipID = CStr(objLMPlantItem.Id)
            End If
        Else
            strEquipID = CStr(Item.EquipmentID)
        End If

        'get the Equipment
        objLMEquipment = Datasource.GetEquipment(strEquipID)
        objLMEquipment.Attributes.BuildAttributesOnDemand = True
        If Not IsDBNull(objLMEquipment.Attributes("ItemTag").Value) Then
            strEquipTag = objLMEquipment.Attributes("ItemTag").Value
        End If


MainLabel:
        If intResponse = MsgBoxResult.Yes Then
            If m_NozzleSeqNo = 0 Then m_NozzleSeqNo = 1
            'reset the starting Nozzle Tag Seq No to 1 if the equipment is different
            If Len(m_PrevNozzleEqID) = 0 Then
                m_PrevNozzleEqID = objLMEquipment.Id
            ElseIf m_PrevNozzleEqID <> objLMEquipment.Id Then
                m_PrevNozzleEqID = objLMEquipment.Id
                'reset the next TagSeqNo to 1 if the equipment is different.
                m_NozzleSeqNo = 1
            End If

            strLocTagSeqNo = CStr(m_NozzleSeqNo + 1)
            m_NozzleSeqNo = CInt(strLocTagSeqNo)
            strItemTag = BuildNozzleTag(colTagValues.Item(CONST_TagPrefixAttributeName), strLocTagSeqNo, colTagValues.Item(CONST_TagSuffixAttributeName))
        End If

        strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Nozzle on Equipment " & strEquipTag & " " & My.Resources.str5007
        m_objLMNozzles = New Llama.LMNozzles

        For Each objLMNozzle In objLMEquipment.Nozzles
            m_objLMNozzles.Add(objLMNozzle.AsLMAItem)
        Next objLMNozzle

        ' The nozzles on the children
        For Each objLMPlantItem In objLMEquipment.ChildPlantItemPlantItems
            Call GetChildNozzles(objLMPlantItem)
        Next objLMPlantItem

        For Each objLMNozzle In m_objLMNozzles
            objLMNozzle.Attributes.BuildAttributesOnDemand = True

            If objLMNozzle.ItemStatusIndex <> 4 Then
                If Not IsDBNull(objLMNozzle.ItemTag) Then
                    If objLMNozzle.ItemTag = strItemTag Then
                        If objLMNozzle.IsBulkItemIndex = CONST_TrueIndex And nBulkItemIndex = CONST_TrueIndex Or objLMNozzle.Id = Item.Id Then
                            UpdateNozzleTag = True
                        Else
                            UpdateNozzleTag = False
                            GoTo ExitDueToDupInActivePlant
                        End If
                    End If
                End If
            End If
        Next objLMNozzle

        If UpdateNozzleTag = True Then
            If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuilt Then
                strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Nozzle on Equipment " & strEquipTag & " " & My.Resources.str5008
                objFilter = Nothing
                objFilter = GetFilter(CONST_NozzleItemName, CONST_ItemTagAttributeName, strItemTag)
                UpdateNozzleTag = CheckForAsBuiltUnique(Datasource, objFilter, CStr(Item.Id), nBulkItemIndex, strItemTag, strEquipID)
            End If
        End If

        If UpdateNozzleTag = True Then
            If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuiltAndProjs Then
                strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Nozzle on Equipment " & strEquipTag & " " & My.Resources.str5011
                objFilter = Nothing
                objFilter = GetFilter(CONST_NozzleItemName, CONST_ItemTagAttributeName, strItemTag)
                UpdateNozzleTag = CheckForGlobalUnique(Datasource, objFilter, CStr(Item.Id), nBulkItemIndex, strItemTag, strEquipID)
            End If
        End If

ExitDueToDupInActivePlant:

        If UpdateNozzleTag = True Then
ExitDueTOTagSeq:
            Item.Attributes.Item(CONST_ItemTagAttributeName).Value = strItemTag
            If Len(strLocTagSeqNo) > 0 Then
                Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value = strLocTagSeqNo
            Else
                Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value = System.DBNull.Value
            End If

            If Not IsDBNull(varValue) Then
                'trim off any spaces in the input value
                varValue = Trim(varValue)
            End If

            If strPropertyName = CONST_TagPrefixAttributeName Then
                Item.Attributes.Item(CONST_TagPrefixAttributeName).Value = varValue
                If Len(colTagValues.Item(CONST_TagSuffixAttributeName)) > 0 Then
                    Item.Attributes.Item(CONST_TagSuffixAttributeName).Value = colTagValues.Item(CONST_TagSuffixAttributeName)
                End If
            ElseIf strPropertyName = CONST_TagSuffixAttributeName Then
                Item.Attributes.Item(CONST_TagSuffixAttributeName).Value = varValue
                If Len(colTagValues.Item(CONST_TagPrefixAttributeName)) > 0 Then
                    Item.Attributes.Item(CONST_TagPrefixAttributeName).Value = colTagValues.Item(CONST_TagPrefixAttributeName)
                End If
            Else
                If Len(colTagValues.Item(CONST_TagPrefixAttributeName)) > 0 Then
                    Item.Attributes.Item(CONST_TagPrefixAttributeName).Value = colTagValues.Item(CONST_TagPrefixAttributeName)
                End If
                If Len(colTagValues.Item(CONST_TagSuffixAttributeName)) > 0 Then
                    Item.Attributes.Item(CONST_TagSuffixAttributeName).Value = colTagValues.Item(CONST_TagSuffixAttributeName)
                End If
            End If
            Item.Commit()
            UpdateNozzleTag = True
            Exit Function
        Else
            If m_isUIEnabled Then
                If intResponse = MsgBoxResult.Yes Then
                    intResponse = MsgBoxResult.Yes
                Else
                    If DoUntilUni = True Then
                        intResponse = MsgBoxResult.Yes
                    Else
                        If strDupMessage = "" Then
                            intResponse = MsgBox(My.Resources.str5000, MsgBoxStyle.YesNo + MsgBoxStyle.SystemModal + MsgBoxStyle.Exclamation, My.Resources.str5001)
                        Else
                            intResponse = MsgBox(strDupMessage & " " & My.Resources.str5009, MsgBoxStyle.YesNo + MsgBoxStyle.SystemModal + MsgBoxStyle.Exclamation, My.Resources.str5001)
                        End If
                    End If
                End If
            Else
                intResponse = MsgBoxResult.Yes
            End If
            If intResponse = MsgBoxResult.Yes Then
                UpdateNozzleTag = True
                GoTo MainLabel
            Else
                UpdateNozzleTag = False
            End If
        End If
        'cleanup
        objLMPlantItem = Nothing
        objChildPlantItem = Nothing
        objLMEquipment = Nothing
        objEquipComp = Nothing
        objLMNozzle = Nothing
        vPartOfID = System.DBNull.Value
        objFilter = Nothing
        Exit Function

errHandler:
        'cleanup
        objLMPlantItem = Nothing
        objChildPlantItem = Nothing
        objLMEquipment = Nothing
        objEquipComp = Nothing
        objLMNozzle = Nothing
        vPartOfID = System.DBNull.Value
        objFilter = Nothing
    End Function


    Private Function UpdateLoopTag(ByRef Datasource As Llama.LMADataSource, ByRef Item As Llama.LMInstrLoop, ByRef varValue As Object, ByRef strPropertyName As String) As Boolean

        Dim colTagValues As Collection
        Dim objFilter As Llama.LMAFilter
        Dim strDupMessage As String
        Dim strItemTag As String
        Dim strLocTagSeqNo As String
        Dim intResponse As Short
        Dim blnUnique As Boolean
        Dim bUniqueStatus As Boolean

        On Error GoTo ErrorHandler

        Item.Attributes.BuildAttributesOnDemand = True
        strDupMessage = ""

        'Use the collections unique index feature to avoid a bunch of if statements to
        'determine what property was passed in
        colTagValues = New Collection

        With colTagValues

            .Add(Trim(VariantToString(varValue)), strPropertyName)

            Select Case strPropertyName

                Case CONST_TagSuffixAttributeName
                    .Add(Trim(VariantToString((Item.LoopFunction))), CONST_LoopFunctionAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)

                Case CONST_TagSequenceNoAttributeName
                    .Add(Trim(VariantToString((Item.LoopFunction))), CONST_LoopFunctionAttributeName)
                    .Add(Trim(VariantToString((Item.TagSuffix))), CONST_TagSuffixAttributeName)

                Case CONST_LoopFunctionAttributeName
                    .Add(Trim(VariantToString((Item.TagSuffix))), CONST_TagSuffixAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)

                Case Else
                    .Add(Trim(VariantToString((Item.TagSuffix))), CONST_TagSuffixAttributeName)
                    .Add(Trim(VariantToString((Item.LoopFunction))), CONST_LoopFunctionAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)
            End Select
        End With

        strLocTagSeqNo = colTagValues.Item(CONST_TagSequenceNoAttributeName)

        If Len(strLocTagSeqNo) = 0 And strPropertyName <> CONST_TagSequenceNoAttributeName Then
            'Get the next tagseqno from Options Manager
            strLocTagSeqNo = GetNextAvailTagSeqNo(Datasource, CONST_NameAttributeName, CONST_InstrLoopNextSeqNoAttributeName)
        End If

        strItemTag = BuildInstrLoopTag(colTagValues.Item(CONST_LoopFunctionAttributeName), strLocTagSeqNo, colTagValues.Item(CONST_TagSuffixAttributeName))
        UpdateLoopTag = True

        If Len(strLocTagSeqNo) > 0 Then
            If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuiltAndProjs Then
                strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Loop " & My.Resources.str5011
                objFilter = GetFilter(CONST_InstrLoopItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_InstrLoopPlantItemIndex))
                'Checking for Unique
                bUniqueStatus = CheckForGlobalUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex), strItemTag, "")
            Else
                If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuilt Then
                    strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Loop " & My.Resources.str5010
                Else
                    strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Loop " & My.Resources.str5007
                End If

                'Check the active plant
                objFilter = GetFilter(CONST_InstrLoopItemName, CONST_ItemTagAttributeName, strItemTag, CONST_ItemStatus, Const_ItemStatusValue)
                'Checking for Unique
                bUniqueStatus = CheckForUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex))

                'Check to see if the Asbuilt plant needs checking too
                If bUniqueStatus = True Then
                    If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuilt Then
                        strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Loop " & My.Resources.str5008
                        objFilter = Nothing
                        objFilter = GetFilter(CONST_InstrLoopItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_InstrLoopPlantItemIndex))
                        'Checking for Unique
                        bUniqueStatus = CheckForAsBuiltUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex), strItemTag, "")
                    End If
                End If
            End If

            If bUniqueStatus = False Then
                If m_isUIEnabled Then
                    intResponse = MsgBox(strDupMessage & " " & My.Resources.str5009, MsgBoxStyle.YesNo + MsgBoxStyle.SystemModal + MsgBoxStyle.Exclamation, My.Resources.str5001)
                Else
                    intResponse = MsgBoxResult.Yes
                End If

                If intResponse = MsgBoxResult.Yes Then

                    blnUnique = False

                    Do While blnUnique = False

                        'Get the next tagseqno from Options Manager
                        strLocTagSeqNo = GetNextAvailTagSeqNo(Datasource, CONST_NameAttributeName, CONST_InstrLoopNextSeqNoAttributeName)

                        strItemTag = BuildInstrLoopTag(colTagValues.Item(CONST_LoopFunctionAttributeName), strLocTagSeqNo, colTagValues.Item(CONST_TagSuffixAttributeName))
                        objFilter = Nothing
                        If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuiltAndProjs Then
                            objFilter = GetFilter(CONST_InstrLoopItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_InstrLoopPlantItemIndex))
                            'Checking for Unique
                            blnUnique = CheckForGlobalUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex), strItemTag, "")
                        Else
                            'Forming a filter
                            objFilter = GetFilter(CONST_InstrLoopItemName, CONST_ItemTagAttributeName, strItemTag, CONST_ItemStatus, Const_ItemStatusValue)
                            'Checking for Unique
                            blnUnique = CheckForUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex))
                            If blnUnique = True Then
                                If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuilt Then
                                    objFilter = Nothing
                                    objFilter = GetFilter(CONST_InstrLoopItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_InstrLoopPlantItemIndex))
                                    blnUnique = CheckForAsBuiltUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex), strItemTag, "")
                                End If
                            End If
                        End If
                    Loop
                Else
                    UpdateLoopTag = False
                End If
            End If
        End If

        If UpdateLoopTag Then

            With Item

                If Len(strLocTagSeqNo) > 0 Then
                    .Attributes.Item(CONST_TagSequenceNoAttributeName).Value = strLocTagSeqNo
                Else
                    .Attributes.Item(CONST_TagSequenceNoAttributeName).Value = System.DBNull.Value
                End If

                If strLocTagSeqNo = "" Then
                    Item.ItemTag = System.DBNull.Value
                Else
                    If Len(strItemTag) > 0 Then
                        Item.ItemTag = strItemTag
                    End If
                    If Len(colTagValues.Item(CONST_TagSuffixAttributeName)) > 0 Then
                        .Attributes.Item(CONST_TagSuffixAttributeName).Value = colTagValues.Item(CONST_TagSuffixAttributeName)
                    End If
                End If

                If Not IsDBNull(varValue) Then
                    'trim off any spaces in the input value
                    varValue = Trim(varValue)
                End If

                .Commit()

            End With

        End If

        Exit Function

ErrorHandler:
        LogError("ItemTag - ItemTagFunc::UpdateLoopTag --> " & Err.Description)
        colTagValues = Nothing
        objFilter = Nothing

    End Function

    'Function to get the next tagseqno from Options Manager
    Private Function GetNextAvailTagSeqNo(ByVal objDataSource As Llama.LMADataSource, ByVal SourceAttributeName As String, ByVal ValueAttribute As String) As String

        On Error GoTo ErrorHandler

        Dim objOptionSettings As New Llama.LMOptionSettings
        Dim objFilter As Llama.LMAFilter

        objOptionSettings = New Llama.LMOptionSettings

        'Forming a filter
        objFilter = GetFilter(CONST_OptionSettingAttributeName, SourceAttributeName, ValueAttribute)

        With objDataSource

            .QueryCache = False

            With objOptionSettings

                .Collect(objDataSource, Filter:=objFilter)
                GetNextAvailTagSeqNo = Trim(.Nth(1).Value)

                With .Nth(1)
                    .Value = CInt(GetNextAvailTagSeqNo) + 1
                    .Commit()
                End With

            End With

            .QueryCache = True

        End With

Cleanup:
        objOptionSettings = Nothing
        objFilter = Nothing

        Exit Function

ErrorHandler:
        LogError("ItemTag - ItemTagFunc::GetNextAvailTagSeqNo --> " & Err.Description)
        Resume Cleanup

    End Function

    'Function that forms a filter
    Private Function GetFilter(ByVal ItemType As String, ByRef SourceAttributeName As String, ByVal ValueAttribute As String, Optional ByRef SourceAttributeName1 As String = "", Optional ByRef ValueAttribute1 As String = "", Optional ByRef SourceAttributeName2 As String = "", Optional ByRef ValueAttribute2 As String = "", Optional ByRef SourceAttributeName3 As String = "", Optional ByRef ValueAttribute3 As String = "") As Llama.LMAFilter

        Dim objCriterion As Llama.LMACriterion

        On Error GoTo ErrorHandler

        GetFilter = New Llama.LMAFilter

        With GetFilter

            .ItemType = ItemType
            .Conjunctive = True

            objCriterion = New Llama.LMACriterion

            With objCriterion
                .Conjunctive = True
                .SourceAttributeName = SourceAttributeName
                .ValueAttribute = ValueAttribute
                .Operator = "="
                .Bind = True
            End With

            .Criteria.Add(objCriterion)

            If Len(SourceAttributeName1) > 1 Then

                objCriterion = Nothing
                objCriterion = New Llama.LMACriterion

                With objCriterion
                    .Conjunctive = True
                    .SourceAttributeName = SourceAttributeName1
                    .ValueAttribute = ValueAttribute1
                    .Operator = "="
                End With

                GetFilter.Criteria.Add(objCriterion)

            End If

            If Len(SourceAttributeName2) > 1 Then

                objCriterion = Nothing
                objCriterion = New Llama.LMACriterion

                With objCriterion
                    .Conjunctive = True
                    .SourceAttributeName = SourceAttributeName2
                    .ValueAttribute = ValueAttribute2
                    .Operator = "="
                End With

                GetFilter.Criteria.Add(objCriterion)

            End If

            If Len(SourceAttributeName3) > 1 Then

                objCriterion = Nothing
                objCriterion = New Llama.LMACriterion

                With objCriterion
                    .Conjunctive = True
                    .SourceAttributeName = SourceAttributeName3
                    .ValueAttribute = ValueAttribute3
                    .Operator = "="
                End With

                GetFilter.Criteria.Add(objCriterion)

            End If

        End With

        Exit Function

ErrorHandler:
        LogError("ItemTag - ItemTagFunc::GetFilter --> " & Err.Description)
        GetFilter = Nothing
        objCriterion = Nothing

    End Function

    'Function to check if itemtags are unique
    Private Function CheckForUnique(ByVal Datasource As Llama.LMADataSource, ByVal objFilter As Llama.LMAFilter, ByVal IDOFITEM As String, ByRef BulkItemIndex As Integer) As Boolean

        Dim objLMPipeRuns As Llama.LMPipeRuns
        Dim objLMDuctRuns As Llama.LMDuctRuns
        Dim objLMInstruments As Llama.LMInstruments
        Dim objLMInstrLoops As Llama.LMInstrLoops
        Dim objLMEquipments As Llama.LMEquipments
        Dim objLMRooms As Llama.LMRooms
        Dim objLMNozzles As Llama.LMNozzles
        Dim objLMSignalRuns As Llama.LMSignalRuns
        On Error GoTo ErrorHandler

        With objFilter
            'Check the Piperun collection
            If .ItemType = CONST_PipeRunItemName Then

                objLMPipeRuns = New Llama.LMPipeRuns

                With objLMPipeRuns

                    .Collect(Datasource, Filter:=objFilter)

                    If .Count > 0 Then
                        .Nth(1).Attributes.BuildAttributesOnDemand = True

                        If .Nth(1).IsBulkItemIndex = CONST_TrueIndex And BulkItemIndex = CONST_TrueIndex Then
                            CheckForUnique = True
                        ElseIf .Count = 1 And CStr(.Nth(1).Id) = CStr(IDOFITEM) Then
                            CheckForUnique = True
                        Else
                            CheckForUnique = False
                        End If

                    Else
                        CheckForUnique = True
                    End If

                End With
                'Check the Ductrun collection
            ElseIf .ItemType = CONST_DuctRunItemName Then

                objLMDuctRuns = New Llama.LMDuctRuns

                With objLMDuctRuns

                    .Collect(Datasource, Filter:=objFilter)

                    If .Count > 0 Then
                        .Nth(1).Attributes.BuildAttributesOnDemand = True

                        If .Nth(1).IsBulkItemIndex = CONST_TrueIndex And BulkItemIndex = CONST_TrueIndex Then
                            CheckForUnique = True
                        ElseIf .Count = 1 And CStr(.Nth(1).Id) = CStr(IDOFITEM) Then
                            CheckForUnique = True
                        Else
                            CheckForUnique = False
                        End If

                    Else
                        CheckForUnique = True
                    End If

                End With

                'Check the Instrument collection
            ElseIf .ItemType = CONST_InstrumentItemName Then

                objLMInstruments = New Llama.LMInstruments

                With objLMInstruments

                    .Collect(Datasource, Filter:=objFilter)

                    If .Count > 0 Then
                        .Nth(1).Attributes.BuildAttributesOnDemand = True

                        If .Nth(1).IsBulkItemIndex = CONST_TrueIndex And BulkItemIndex = CONST_TrueIndex Then
                            CheckForUnique = True
                        ElseIf .Count = 1 And CStr(.Nth(1).Id) = CStr(IDOFITEM) Then
                            CheckForUnique = True
                        Else
                            CheckForUnique = False
                        End If

                    Else
                        CheckForUnique = True
                    End If

                End With

                'Check the InstrumentLoop collection
            ElseIf .ItemType = CONST_InstrLoopItemName Then

                objLMInstrLoops = New Llama.LMInstrLoops

                With objLMInstrLoops

                    .Collect(Datasource, Filter:=objFilter)

                    If .Count > 0 Then

                        .Nth(1).Attributes.BuildAttributesOnDemand = True
                        If .Nth(1).IsBulkItemIndex = CONST_TrueIndex And BulkItemIndex = CONST_TrueIndex Then
                            CheckForUnique = True
                        ElseIf .Count = 1 And CStr(objLMInstrLoops.Nth(1).Id) = CStr(IDOFITEM) Then
                            CheckForUnique = True
                        Else
                            CheckForUnique = False
                        End If

                    Else
                        CheckForUnique = True
                    End If

                End With

                'Check the Equipment collection
            ElseIf .ItemType = CONST_EquipmentItemName Then

                objLMEquipments = New Llama.LMEquipments

                With objLMEquipments

                    .Collect(Datasource, Filter:=objFilter)

                    If .Count > 0 Then

                        .Nth(1).Attributes.BuildAttributesOnDemand = True
                        If .Nth(1).IsBulkItemIndex = CONST_TrueIndex And BulkItemIndex = CONST_TrueIndex Then
                            CheckForUnique = True
                        ElseIf .Count = 1 And CStr(.Nth(1).Id) = CStr(IDOFITEM) Then
                            CheckForUnique = True
                        Else
                            CheckForUnique = False
                        End If

                    Else
                        CheckForUnique = True
                    End If

                End With

                'Check the Room collection
            ElseIf .ItemType = CONST_RoomItemName Then

                objLMRooms = New Llama.LMRooms

                With objLMRooms

                    .Collect(Datasource, Filter:=objFilter)

                    If .Count > 0 Then

                        .Nth(1).Attributes.BuildAttributesOnDemand = True
                        If .Nth(1).IsBulkItemIndex = CONST_TrueIndex And BulkItemIndex = CONST_TrueIndex Then
                            CheckForUnique = True
                        ElseIf .Count = 1 And CStr(.Nth(1).Id) = CStr(IDOFITEM) Then
                            CheckForUnique = True
                        Else
                            CheckForUnique = False
                        End If

                    Else
                        CheckForUnique = True
                    End If

                End With

                'Checking in the Nozzles collection
            ElseIf .ItemType = CONST_NozzleItemName Then

                objLMNozzles = New Llama.LMNozzles

                With objLMNozzles

                    .Collect(Datasource, Filter:=objFilter)

                    If .Count > 0 Then

                        .Nth(1).Attributes.BuildAttributesOnDemand = True
                        If .Nth(1).IsBulkItemIndex = CONST_TrueIndex And BulkItemIndex = CONST_TrueIndex Then
                            CheckForUnique = True
                        ElseIf .Count = 1 And CStr(.Nth(1).Id) = CStr(IDOFITEM) Then
                            CheckForUnique = True
                        Else
                            CheckForUnique = False
                        End If

                    Else
                        CheckForUnique = True
                    End If

                End With
            ElseIf .ItemType = Const_SignalRunItemName Then

                objLMSignalRuns = New Llama.LMSignalRuns

                With objLMSignalRuns

                    .Collect(Datasource, Filter:=objFilter)

                    If .Count > 0 Then

                        .Nth(1).Attributes.BuildAttributesOnDemand = True
                        If .Nth(1).IsBulkItemIndex = CONST_TrueIndex And BulkItemIndex = CONST_TrueIndex Then
                            CheckForUnique = True
                        ElseIf .Count = 1 And CStr(.Nth(1).Id) = CStr(IDOFITEM) Then
                            CheckForUnique = True
                        Else
                            CheckForUnique = False
                        End If

                    Else
                        CheckForUnique = True
                    End If

                End With

            End If

        End With

        Exit Function

ErrorHandler:
        LogError("ItemTag - ItemTagFunc::CheckForUnique --> " & Err.Description)
        objLMPipeRuns = Nothing
        objLMInstruments = Nothing
        objLMInstrLoops = Nothing
        objLMEquipments = Nothing
        objLMNozzles = Nothing

    End Function



    Private Function GetUnit(ByRef Item As Llama.LMPlantItem) As Llama.LMUnit

        Dim objPlantGroup As Llama.LMPlantGroup

        On Error GoTo ErrorHandler
        GetUnit = Nothing
        'Get the Unit
        objPlantGroup = Item.PlantGroupObject
        If Not (objPlantGroup Is Nothing) Then
            With objPlantGroup

                If .PlantGroupTypeIndex = 2 Then
                    GetUnit = .DataSource.GetUnit(.Id)
                End If

            End With
        End If
        Exit Function

ErrorHandler:
        LogError("ItemTag - ItemTagFunc::GetUnit --> " & Err.Description)
        GetUnit = Nothing

    End Function

    Private Function VariantToString(ByRef varValue As Object) As String

        VariantToString = IIf(IsDBNull(varValue), "", varValue)

    End Function

    Private Function IsNozzleTagUnique(ByRef objLMPlantItem As Llama.LMPlantItem, ByRef value As Object, ByVal BulkItemIndex As Integer) As Boolean

        Dim objLMChildPlantItem As Llama.LMPlantItem
        Dim bIsNozzleTagUnique As Boolean
        Dim objLMNozzle As Llama.LMNozzle

        On Error GoTo errHandler

        bIsNozzleTagUnique = True

        If objLMPlantItem.ItemTypeName = "Nozzle" Then
            objLMNozzle = objLMPlantItem.DataSource.GetNozzle(CStr(objLMPlantItem.Id))
            If (Not IsDBNull(value)) And (Not IsDBNull(objLMNozzle.ItemTag)) Then
                If objLMNozzle.ItemTag = value Then
                    If objLMNozzle.IsBulkItemIndex = CONST_TrueIndex And BulkItemIndex = CONST_TrueIndex Then
                        bIsNozzleTagUnique = True
                    Else
                        bIsNozzleTagUnique = False
                        IsNozzleTagUnique = False
                        Exit Function
                    End If
                    If CInt(objLMNozzle.Attributes.Item(CONST_TagSequenceNoAttributeName).Value) > m_NozzleSeqNo Then
                        m_NozzleSeqNo = CInt(objLMNozzle.Attributes.Item(CONST_TagSequenceNoAttributeName).Value)
                    End If
                End If
            End If
        Else
            For Each objLMChildPlantItem In objLMPlantItem.ChildPlantItemPlantItems
                bIsNozzleTagUnique = IsNozzleTagUnique(objLMChildPlantItem, value, BulkItemIndex)
                If bIsNozzleTagUnique = False Then
                    IsNozzleTagUnique = False
                    Exit Function
                End If
            Next objLMChildPlantItem
        End If

        IsNozzleTagUnique = bIsNozzleTagUnique

        'cleanup
        objLMChildPlantItem = Nothing

        Exit Function

errHandler:
        'cleanup
        objLMChildPlantItem = Nothing
        IsNozzleTagUnique = False
    End Function


    Private Function UpdateSignalRunTag(ByRef Datasource As Llama.LMADataSource, ByRef Item As Llama.LMSignalRun, ByRef varValue As Object, ByRef strPropertyName As String) As Boolean

        Dim colTagValues As Collection
        Dim objFilter As Llama.LMAFilter
        Dim strDupMessage As String
        Dim strItemTag As String
        Dim strLocTagSeqNo As String
        Dim intResponse As Short
        Dim bStatus As Boolean

        On Error Resume Next

        Item.Attributes.BuildAttributesOnDemand = True
        strDupMessage = ""

        'Use the collections unique index feature to avoid a bunch of if statements to
        'determine what property was passed in
        colTagValues = New Collection

        With colTagValues

            .Add(Trim(VariantToString(varValue)), strPropertyName)

            Select Case strPropertyName

                Case CONST_TagSuffixAttributeName
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)

                Case CONST_TagSequenceNoAttributeName
                    .Add(Trim(VariantToString((Item.TagSuffix))), CONST_TagSuffixAttributeName)
                Case Else
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSuffixAttributeName).Value))), CONST_TagSuffixAttributeName)
            End Select

        End With

        strLocTagSeqNo = colTagValues.Item(CONST_TagSequenceNoAttributeName)

        strItemTag = BuildSignalRunTag(strLocTagSeqNo, colTagValues.Item(CONST_TagSuffixAttributeName))

        UpdateSignalRunTag = True

        If Len(strLocTagSeqNo) > 0 And m_isUIEnabled = True Then
            If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuiltAndProjs Then
                strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Signal Run " & My.Resources.str5011
                'Forming a filter
                'note - CONST_InstrumentPlantItemIndex is really = 5 (SignalRun)
                If Item.IsBulkItemIndex = CONST_TrueIndex Then
                    objFilter = GetFilter(Const_SignalRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_InstrumentPlantItemIndex), CONST_IsBulkItemAttributeName, CStr(CONST_FalseIndex))
                Else
                    objFilter = GetFilter(Const_SignalRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_InstrumentPlantItemIndex))
                End If

                'Checking if the itemtag is unique
                bStatus = CheckForGlobalUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex), strItemTag, "")
                If bStatus = False Then
                    If Item.Id <> m_lngID Then
                        m_lngID = Item.Id
                        intResponse = MsgBox(strDupMessage, MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, My.Resources.str5001)
                    End If
                    'Rami Weiss 31/05/10 Inform the user bat accept the changes - It is only OK button message. Task TK-19524
                    UpdateSignalRunTag = True
                End If
            Else
                If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuilt Then
                    'strDupMessage = "Duplicate Tag " & strItemTag & " found in active project."
                    strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Signal Run " & My.Resources.str5010
                Else
                    'strDupMessage = "Duplicate Tag " & strItemTag & " found in active plant."
                    strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Signal Run " & My.Resources.str5007
                End If
                'check local project - same as always
                objFilter = GetFilter(Const_SignalRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_ItemStatus, Const_ItemStatusValue)

                'Checking if the itemtag is unique in the local project
                If Not CheckForUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex)) Then
                    If Item.Id <> m_lngID Then
                        m_lngID = Item.Id
                        intResponse = MsgBox(strDupMessage, MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, My.Resources.str5001)
                    End If
                    'Rami Weiss 31/05/10 Inform the user bat accept the changes - It is only OK button message. Task TK-19524
                    UpdateSignalRunTag = True
                Else
                    'its unique in the local project but check further into AsBuilt if required
                    If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuilt Then
                        strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Signal Run " & My.Resources.str5008
                        objFilter = GetFilter(Const_SignalRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_InstrumentPlantItemIndex))

                        'Checking if the itemtag is unique
                        bStatus = CheckForAsBuiltUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex), strItemTag, "")
                        If bStatus = False Then
                            If Item.Id <> m_lngID Then
                                m_lngID = Item.Id
                                intResponse = MsgBox(strDupMessage, MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, My.Resources.str5001)
                            End If
                            'Rami Weiss 31/05/10 Inform the user bat accept the changes - It is only OK button message. Task TK-19524
                            UpdateSignalRunTag = True
                        End If
                    End If
                End If
            End If
        End If

        'Update the values
        With Item

            If Len(strLocTagSeqNo) > 0 Then
                .Attributes.Item(CONST_TagSequenceNoAttributeName).Value = strLocTagSeqNo
            Else
                .Attributes.Item(CONST_TagSequenceNoAttributeName).Value = System.DBNull.Value
            End If

            If strLocTagSeqNo = "" Then
                .ItemTag = System.DBNull.Value
            Else
                If Len(colTagValues.Item(CONST_TagSuffixAttributeName)) > 0 Then
                    .TagSuffix = colTagValues.Item(CONST_TagSuffixAttributeName)
                End If

                If Len(strItemTag) > 0 Then
                    .ItemTag = strItemTag
                End If
            End If

            If strLocTagSeqNo = "" Then
                If strPropertyName = CONST_TagSuffixAttributeName Then
                    If Len(colTagValues.Item(CONST_TagSuffixAttributeName)) > 0 Then
                        .Attributes.Item(CONST_TagSuffixAttributeName).Value = colTagValues.Item(CONST_TagSuffixAttributeName)
                    End If
                End If
            End If

            If Not IsDBNull(varValue) Then
                'trim off any spaces in the input value
                varValue = Trim(varValue)
            End If

            .Commit()

        End With

        Err.Clear()
        Exit Function
ErrorHandler:
        UpdateSignalRunTag = False
    End Function

    Private Function GetFilter2(ByVal ItemType As String, ByRef SourceAttributeName As String, ByVal OperatorAttribute As String, ByVal ValueAttribute As String, Optional ByRef SourceAttributeName1 As String = "", Optional ByRef OperatorAttribute1 As String = "", Optional ByRef ValueAttribute1 As String = "", Optional ByRef SourceAttributeName2 As String = "", Optional ByRef OperatorAttribute2 As String = "", Optional ByRef ValueAttribute2 As String = "", Optional ByRef SourceAttributeName3 As String = "", Optional ByRef OperatorAttribute3 As String = "", Optional ByRef ValueAttribute3 As String = "") As Llama.LMAFilter

        Dim objCriterion As Llama.LMACriterion

        On Error GoTo ErrorHandler

        GetFilter2 = New Llama.LMAFilter

        With GetFilter2

            .ItemType = ItemType
            .Conjunctive = True

            objCriterion = New Llama.LMACriterion

            With objCriterion
                .Conjunctive = True
                .SourceAttributeName = SourceAttributeName
                .ValueAttribute = ValueAttribute
                .Operator = OperatorAttribute
                .Bind = True
            End With

            .Criteria.Add(objCriterion)

            If Len(SourceAttributeName1) > 1 Then

                objCriterion = Nothing
                objCriterion = New Llama.LMACriterion

                With objCriterion
                    .Conjunctive = True
                    .SourceAttributeName = SourceAttributeName1
                    .ValueAttribute = ValueAttribute1
                    .Operator = OperatorAttribute1
                End With

                GetFilter2.Criteria.Add(objCriterion)

            End If

            If Len(SourceAttributeName2) > 1 Then

                objCriterion = Nothing
                objCriterion = New Llama.LMACriterion

                With objCriterion
                    .Conjunctive = True
                    .SourceAttributeName = SourceAttributeName2
                    .ValueAttribute = ValueAttribute2
                    .Operator = OperatorAttribute2
                End With

                GetFilter2.Criteria.Add(objCriterion)

            End If

            If Len(SourceAttributeName3) > 1 Then

                objCriterion = Nothing
                objCriterion = New Llama.LMACriterion

                With objCriterion
                    .Conjunctive = True
                    .SourceAttributeName = SourceAttributeName3
                    .ValueAttribute = ValueAttribute3
                    .Operator = OperatorAttribute3
                End With

                GetFilter2.Criteria.Add(objCriterion)
            End If

        End With

        Exit Function

ErrorHandler:
        LogError("ItemTag - ItemTagFunc::GetFilter2 --> " & Err.Description)
        GetFilter2 = Nothing
        objCriterion = Nothing

    End Function

    'Function to check if itemtags are unique
    Private Function CheckForUniqueUsingViews(ByRef strTheView As String, ByRef Datasource As Llama.LMADataSource, ByRef objFilter As Llama.LMAFilter, ByRef IDOFITEM As String, ByRef BulkItemIndex As Integer, ByRef strItemTag As String, ByRef strParentID As String) As Boolean

        Dim bStatus As Boolean
        Dim lngItemStatusIndex As Integer
        Dim lngPlantItemGroupIndex As Integer
        Dim colIDs As VBA.Collection
        Dim colIsBulkItemIndexes As VBA.Collection
        Dim colPlantItemTypeIndexes As VBA.Collection
        Dim colTags As VBA.Collection
        Dim status As Integer
        Dim strSPID As Object
        Dim strIsBulkItemIndex As String
        Dim vBulkItemStatus As Object
        Dim objVB6Convert As New InteropigrVB6Helper.Helper

        On Error GoTo ErrorHandler

        'If certain item types does not require uniqueness check then this function shouldn't
        'even be called for those item types

        CheckForUniqueUsingViews = True

        With objFilter

            'Check the Piperuns for duplicates
            If .ItemType = CONST_PipeRunItemName Then
                CheckForUniqueUsingViews = True

                colIDs = objVB6Convert.GetEmptyCollection()
                status = Datasource.GetTagDataFromView(strTheView, objFilter, colIDs)

                If status <> 0 Then
                    GoTo ErrorHandler
                Else
                    If colIDs.Count() > 0 Then
                        For Each strSPID In colIDs
                            If strSPID <> IDOFITEM Then
                                'check to see if the piperun with the duplicate tag has an active status
                                status = Datasource.GetModelItemStatusFromView(strTheView, CStr(strSPID), lngItemStatusIndex)
                                If lngItemStatusIndex = CDbl(Const_ItemStatusValue) Then
                                    CheckForUniqueUsingViews = False
                                    GoTo Cleanup
                                End If
                            End If
                        Next strSPID
                    End If
                End If

                'Check the Instruments
            ElseIf .ItemType = CONST_InstrumentItemName Then
                CheckForUniqueUsingViews = True

                colIDs = objVB6Convert.GetEmptyCollection()
                status = Datasource.GetTagDataFromView(strTheView, objFilter, colIDs)
                If status <> 0 Then
                    GoTo ErrorHandler
                Else
                    If colIDs.Count() > 0 Then
                        For Each strSPID In colIDs
                            If strSPID <> IDOFITEM Then
                                'check to see if the instrument with the duplicate tag has an active status
                                status = Datasource.GetModelItemStatusFromView(strTheView, CStr(strSPID), lngItemStatusIndex)
                                If lngItemStatusIndex = CDbl(Const_ItemStatusValue) Then
                                    CheckForUniqueUsingViews = False
                                    GoTo Cleanup
                                End If
                            End If
                        Next strSPID
                    End If
                End If

                'Check the InstrumentLoops for duplicates
            ElseIf .ItemType = CONST_InstrLoopItemName Then
                CheckForUniqueUsingViews = True

                colIDs = objVB6Convert.GetEmptyCollection()
                status = Datasource.GetTagDataFromView(strTheView, objFilter, colIDs)
                If status <> 0 Then
                    GoTo ErrorHandler
                Else
                    If colIDs.Count() > 0 Then
                        For Each strSPID In colIDs
                            If strSPID <> IDOFITEM Then
                                'check to see if the plantitemgroup with the duplicate tag has an active status
                                status = Datasource.GetModelItemStatusFromView(strTheView, CStr(strSPID), lngItemStatusIndex)
                                If lngItemStatusIndex = CDbl(Const_ItemStatusValue) Then
                                    'this is an active plantitemgroup.  See if it is a loop
                                    status = Datasource.GetPlantItemGroupTypeFromView(strTheView, CStr(strSPID), lngPlantItemGroupIndex)
                                    If lngPlantItemGroupIndex = 6 Then
                                        CheckForUniqueUsingViews = False
                                        GoTo Cleanup
                                    End If
                                End If
                            End If
                        Next strSPID
                    End If
                End If

                'Check the Equipment collection
            ElseIf .ItemType = CONST_EquipmentItemName Then
                CheckForUniqueUsingViews = True

                colIDs = objVB6Convert.GetEmptyCollection()
                status = Datasource.GetTagDataFromView(strTheView, objFilter, colIDs)
                If status <> 0 Then
                    GoTo ErrorHandler
                Else
                    If colIDs.Count() > 0 Then
                        For Each strSPID In colIDs
                            If strSPID <> IDOFITEM Then
                                'check to see if the instrument with the duplicate tag has an active status
                                status = Datasource.GetModelItemStatusFromView(strTheView, CStr(strSPID), lngItemStatusIndex)
                                If lngItemStatusIndex = CDbl(Const_ItemStatusValue) Then
                                    CheckForUniqueUsingViews = False
                                    GoTo Cleanup
                                End If

                            End If
                        Next strSPID
                    End If
                End If

                'Checking in the Nozzles view
            ElseIf .ItemType = CONST_NozzleItemName Then
                CheckForUniqueUsingViews = True
                'this checks the nozzle that have a direct relationship to equipment
                bStatus = EqNozTagIsGlobalUnique(strTheView, Datasource, IDOFITEM, strItemTag, BulkItemIndex, strParentID)
                If bStatus = False Then
                    CheckForUniqueUsingViews = False
                    GoTo Cleanup
                End If
                'this checks the nozzle may be part of equipment components that are children of the main equipment
                'its not likely that these will exist but check anyway
                bStatus = ChildEqNozTagIsGlobalUnique(strTheView, Datasource, IDOFITEM, strItemTag, BulkItemIndex, strParentID)
                If bStatus = False Then
                    CheckForUniqueUsingViews = False
                    GoTo Cleanup
                End If
            ElseIf .ItemType = Const_SignalRunItemName Then
                CheckForUniqueUsingViews = True

                colIDs = objVB6Convert.GetEmptyCollection()
                status = Datasource.GetTagDataFromView(strTheView, objFilter, colIDs)
                If status <> 0 Then
                    GoTo ErrorHandler
                Else
                    If colIDs.Count() > 0 Then
                        For Each strSPID In colIDs
                            If strSPID <> IDOFITEM Then
                                'check to see if the instrument with the duplicate tag has an active status
                                status = Datasource.GetModelItemStatusFromView(strTheView, CStr(strSPID), lngItemStatusIndex)
                                If lngItemStatusIndex = CDbl(Const_ItemStatusValue) Then
                                    CheckForUniqueUsingViews = False
                                    GoTo Cleanup
                                End If
                            End If
                        Next strSPID
                    End If
                End If
            End If

        End With

Cleanup:

        colIDs = Nothing
        colPlantItemTypeIndexes = Nothing
        colTags = Nothing
        Exit Function

ErrorHandler:
        CheckForUniqueUsingViews = True
        colIDs = Nothing
        colPlantItemTypeIndexes = Nothing
        colTags = Nothing
        LogError("ItemTag - ItemTagFunc::CheckForUniqueUsingViews --> " & Err.Description)


    End Function


    Private Function ChildEqNozTagIsGlobalUnique(ByRef strTheView As String, ByRef Datasource As Llama.LMADataSource, ByRef strTheNozzleID As String, ByRef strTheNozzleTag As String, ByRef lngTheNozzleBulkItemIndex As Integer, ByRef strParentID As String) As Boolean
        'This function will use the global views for item tag uniqueness check in a project environment.
        'This function will get the nozzles on the equipment components and test the tags
        'This function can not be used in a satellite plant due to not being able to access data at the host
        Dim bIsUnique As Boolean
        Dim colChildIDs As VBA.Collection
        Dim colTags As VBA.Collection
        Dim colPlantItemTypes As VBA.Collection
        Dim colIsBulkItems As VBA.Collection
        Dim colPartOfIDs As VBA.Collection
        Dim lngBulkItemIndex As Integer
        Dim lngItemStatusIndex As Integer
        Dim objFilter As Llama.LMAFilter
        Dim status As Integer
        Dim strItemTag As String
        Dim strErrorMessage As String
        Dim vChildID As Object
        Dim objVB6Convert As New InteropigrVB6Helper.Helper

        On Error GoTo errHandler

        ChildEqNozTagIsGlobalUnique = True
        bIsUnique = True
        strErrorMessage = ""

        objFilter = GetFilter(CONST_EquipmentItemName, CONST_PartOfIDAttributeName, strParentID)
        colChildIDs = objVB6Convert.GetEmptyCollection()
        colTags = objVB6Convert.GetEmptyCollection()
        colPlantItemTypes = objVB6Convert.GetEmptyCollection()

        colIsBulkItems = objVB6Convert.GetEmptyCollection()
        status = Datasource.GetTagDataFromView(strTheView, objFilter, colChildIDs, colTags, colPlantItemTypes, , colIsBulkItems)
        If status <> 0 Then
            strErrorMessage = "Error getting child items for EquipmentID=" & strParentID & " using global tag view"
            GoTo errHandler
        End If
        For Each vChildID In colChildIDs
            If CInt(colPlantItemTypes.Item(CStr(vChildID)) = CONST_NozzlePlantItemIndex) Then
                If CStr(vChildID) <> strTheNozzleID Then
                    'Found another nozzle related to the parent
                    'Is the tag duplicated
                    strItemTag = ""
                    On Error Resume Next
                    strItemTag = colTags.Item(CStr(vChildID))
                    On Error GoTo errHandler
                    If strItemTag = "" Then
                        GoTo nextChild
                    Else
                        If strItemTag = strTheNozzleTag Then
                            'Found a duplicate, ignore it if both are bulk items
                            lngBulkItemIndex = colIsBulkItems.Item(CStr(vChildID))
                            If CInt(lngBulkItemIndex) = CONST_TrueIndex And lngTheNozzleBulkItemIndex = CONST_TrueIndex Then
                                GoTo nextChild
                            End If
                            'Is the Item Active
                            status = Datasource.GetModelItemStatusFromView(strTheView, CStr(vChildID), lngItemStatusIndex)
                            If lngItemStatusIndex <> 4 Then '4 = Delete Pending
                                bIsUnique = False
                                GoTo Cleanup
                            End If
                        End If
                    End If
                End If
            Else
                bIsUnique = ChildEqNozTagIsGlobalUnique(strTheView, Datasource, strTheNozzleID, strTheNozzleTag, lngTheNozzleBulkItemIndex, CStr(vChildID))
                If bIsUnique = False Then
                    GoTo Cleanup
                End If
            End If
nextChild:

        Next vChildID

Cleanup:
        ChildEqNozTagIsGlobalUnique = bIsUnique
        colChildIDs = Nothing
        objFilter = Nothing
        colChildIDs = Nothing
        colTags = Nothing
        colPlantItemTypes = Nothing
        colIsBulkItems = Nothing

        Exit Function
errHandler:
        bIsUnique = False
        LogError("ItemTag - ChildEqNozTagIsGlobalUnique --> " & Err.Description & " " & strErrorMessage)
        Resume Cleanup


    End Function

    Private Function EqNozTagIsGlobalUnique(ByRef strTheView As String, ByRef Datasource As Llama.LMADataSource, ByRef strTheNozzleID As String, ByRef strTheNozzleTag As String, ByRef lngTheNozzleBulkItemIndex As Integer, ByRef strEquipmentID As String) As Boolean
        'This function will use the global views for item tag uniqueness check in a project environment.
        'This function will get the nozzles on equipment using the SP_EquipmentID and test the tags
        'This function can not be used in a satellite plant due to not being able to access data at the host
        Dim bIsUnique As Boolean
        Dim colDupNozTagIDs As VBA.Collection
        Dim colNozzleIDs As VBA.Collection
        Dim lngItemStatusIndex As Integer
        Dim objFilter As Llama.LMAFilter
        Dim status As Integer
        Dim strErrorMessage As String
        Dim vChildID As Object
        Dim vNozzleID As Object
        Dim objVB6Convert As New InteropigrVB6Helper.Helper

        On Error GoTo errHandler

        EqNozTagIsGlobalUnique = True
        If IsDBNull(strTheNozzleTag) Then
            Exit Function
        End If
        strErrorMessage = ""
        bIsUnique = True

        'Get the nozzles on the parent equipment from the view "GlobalNozzleList"
        colNozzleIDs = objVB6Convert.GetEmptyCollection()
        status = Datasource.GetNozzlesForEquipFromView(strTheView, strEquipmentID, colNozzleIDs)
        If status <> 0 Then
            strErrorMessage = "Error getting nozzles for EquipmentID=" & strEquipmentID & " using global nozzle view"
            GoTo errHandler
        End If

        If colNozzleIDs.Count() > 0 Then
            'check each nozzle to see if the tag is a duplicate
            For Each vNozzleID In colNozzleIDs
                If (CStr(vNozzleID) <> strTheNozzleID) Then
                    objFilter = Nothing
                    If lngTheNozzleBulkItemIndex = CONST_TrueIndex Then
                        objFilter = GetFilter(CONST_NozzleItemName, CONST_SP_IDAttributeName, CStr(vNozzleID), CONST_ItemTagAttributeName, strTheNozzleTag, CONST_IsBulkItemAttributeName, CStr(CONST_FalseIndex))
                    Else
                        objFilter = GetFilter(CONST_NozzleItemName, CONST_SP_IDAttributeName, CStr(vNozzleID), CONST_ItemTagAttributeName, strTheNozzleTag)
                    End If
                    colDupNozTagIDs = objVB6Convert.GetEmptyCollection()
                    'This will query the GlobalTagList view to get the Nozzle with the same tag as the input nozzle
                    status = Datasource.GetTagDataFromView(strTheView, objFilter, colDupNozTagIDs)
                    If status <> 0 Then
                        strErrorMessage = "Error getting nozzle data from global tag view for nozzle id=" & CStr(vNozzleID)
                        GoTo errHandler
                    End If
                    If colDupNozTagIDs.Count() > 0 Then
                        'found a duplicate, is it not delete pending
                        status = Datasource.GetModelItemStatusFromView(strTheView, CStr(vNozzleID), lngItemStatusIndex)
                        If status <> 0 Then
                            strErrorMessage = "Error Getting ModelItemStatus from global view for nozzle id=" & CStr(vNozzleID)
                            GoTo errHandler
                        End If
                        If lngItemStatusIndex <> 4 Then '4 = Delete Pending
                            bIsUnique = False
                            GoTo Cleanup
                        End If
                    End If
                    colDupNozTagIDs = Nothing
                End If
            Next vNozzleID
        End If

Cleanup:
        EqNozTagIsGlobalUnique = bIsUnique
        colDupNozTagIDs = Nothing
        colNozzleIDs = Nothing
        objFilter = Nothing

        Exit Function
errHandler:
        bIsUnique = False
        LogError("ItemTag - EqNozTagIsGlobalUnique --> " & Err.Description & " " & strErrorMessage)
        Resume Cleanup
    End Function

    Private Function CheckForAsBuiltUnique(ByRef Datasource As Llama.LMADataSource, ByRef objFilter As Llama.LMAFilter, ByRef IDOFITEM As String, ByRef BulkItemIndex As Integer, ByRef strItemTag As String, ByRef strParentID As String) As Boolean
        On Error GoTo errHandler

        CheckForAsBuiltUnique = CheckForUniqueUsingViews("AsBuilt", Datasource, objFilter, IDOFITEM, BulkItemIndex, strItemTag, strParentID)
        Exit Function
errHandler:
    End Function

    Private Function CheckForGlobalUnique(ByRef Datasource As Llama.LMADataSource, ByRef objFilter As Llama.LMAFilter, ByRef IDOFITEM As String, ByRef BulkItemIndex As Integer, ByRef strItemTag As String, ByRef strParentID As String) As Boolean
        On Error GoTo errHandler

        CheckForGlobalUnique = CheckForUniqueUsingViews("Global", Datasource, objFilter, IDOFITEM, BulkItemIndex, strItemTag, strParentID)
        Exit Function
errHandler:
    End Function

    Private Function BuildEquipmentTag(ByVal strPrefix As String, ByVal strTagSeqNo As String, ByVal strSuffix As String) As String

        BuildEquipmentTag = strPrefix & "-" & strTagSeqNo & strSuffix

    End Function



    Private Function BuildInstrumentTag(ByVal strMeasuredVarCode As String, ByVal strInstrTypeMod As String, ByVal strInstrTagSeqNo As String, ByVal strLoopTagSuffix As String, ByVal strInstrTagSuffix As String) As String

        BuildInstrumentTag = strMeasuredVarCode & strInstrTypeMod & "-" & strInstrTagSeqNo & strLoopTagSuffix & strInstrTagSuffix

    End Function

    Private Function BuildInstrLoopTag(ByVal strLoopFunc As String, ByVal strTagSeqNo As String, ByVal strSuffix As String) As String

        BuildInstrLoopTag = strLoopFunc & "-" & strTagSeqNo & strSuffix

    End Function

    Private Function BuildNozzleTag(ByVal strPrefix As String, ByVal strTagSeqNo As String, ByVal strSuffix As String) As String

        BuildNozzleTag = strPrefix & strTagSeqNo & strSuffix

    End Function

    Private Function BuildPipeRunTag(ByVal strUnitName As String, ByVal strTagSeqNo As String, ByVal strSuffix As String, ByVal strFluidCode As String) As String

        BuildPipeRunTag = strUnitName & strTagSeqNo & strSuffix
        If Len(strFluidCode) > 0 Then
            BuildPipeRunTag = BuildPipeRunTag & "-" & strFluidCode
        End If

    End Function

    Private Function BuildDuctRunTag(ByVal strUnitName As String, ByVal strTagSeqNo As String, ByVal strSuffix As String, ByVal strFluidCode As String) As String

        BuildDuctRunTag = strUnitName & strTagSeqNo & strSuffix
        If Len(strFluidCode) > 0 Then
            BuildDuctRunTag = BuildDuctRunTag & "-" & strFluidCode
        End If

    End Function

    Private Function BuildSignalRunTag(ByVal strTagSeqNo As String, ByVal strSuffix As String) As String

        BuildSignalRunTag = strTagSeqNo & strSuffix

    End Function

    Private Function UpdateRoomTag(ByRef Datasource As Llama.LMADataSource, ByRef Item As Llama.LMRoom, ByRef varValue As Object, ByRef strPropertyName As String) As Boolean

        Dim colTagValues As Collection
        Dim objFilter As Llama.LMAFilter
        Dim strDupMessage As String
        Dim strItemTag As String
        Dim strLocTagSeqNo As String
        Dim intResponse As Short
        Dim blnUnique As Boolean
        Dim bUniqueStatus As Boolean

        On Error GoTo ErrorHandler

        Item.Attributes.BuildAttributesOnDemand = True
        strDupMessage = ""

        'Use the collections unique index feature to avoid a bunch of if statements
        'to determine what property was passed in
        colTagValues = New Collection

        With colTagValues
            'Trim off any white spaces in the properties that make up the tag.
            .Add(Trim(VariantToString(varValue)), strPropertyName)

            Select Case strPropertyName

                Case CONST_TagSuffixAttributeName
                    .Add(Trim(VariantToString((Item.TagPrefix))), CONST_TagPrefixAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)

                Case CONST_TagPrefixAttributeName
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)
                    .Add(Trim(VariantToString((Item.TagSuffix))), CONST_TagSuffixAttributeName)

                Case CONST_TagSequenceNoAttributeName
                    .Add(Trim(VariantToString((Item.TagPrefix))), CONST_TagPrefixAttributeName)
                    .Add(Trim(VariantToString((Item.TagSuffix))), CONST_TagSuffixAttributeName)

                Case CONST_ItemTagAttributeName
                    .Add(Trim(VariantToString((Item.TagPrefix))), CONST_TagPrefixAttributeName)
                    .Add(Trim(VariantToString((Item.TagSuffix))), CONST_TagSuffixAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)
            End Select

        End With

        strLocTagSeqNo = colTagValues.Item(CONST_TagSequenceNoAttributeName)

        If Len(strLocTagSeqNo) < 1 And strPropertyName <> CONST_TagSequenceNoAttributeName And (Trim(VariantToString(varValue)) <> "" Or strPropertyName = CONST_ItemTagAttributeName) Then

            'Get the next tagseqno from Options Manager
            strLocTagSeqNo = GetNextAvailTagSeqNo(Datasource, CONST_NameAttributeName, CONST_RoomNextSeqNoAttributeName)

            strItemTag = BuildRoomTag(colTagValues.Item(CONST_TagPrefixAttributeName), strLocTagSeqNo, colTagValues.Item(CONST_TagSuffixAttributeName))

            bUniqueStatus = CheckForRoomUnique(Datasource, Item, strItemTag, strDupMessage)

            'Check for uniqueness of the ItemTag
            If bUniqueStatus = False Then
                blnUnique = False

                Do While blnUnique = False

                    'Get the next tagseqno from Options Manager
                    strLocTagSeqNo = GetNextAvailTagSeqNo(Datasource, CONST_NameAttributeName, CONST_RoomNextSeqNoAttributeName)

                    strItemTag = BuildRoomTag(colTagValues.Item(CONST_TagPrefixAttributeName), strLocTagSeqNo, colTagValues.Item(CONST_TagSuffixAttributeName))

                    blnUnique = CheckForRoomUnique(Datasource, Item, strItemTag, strDupMessage)
                Loop
            End If
        End If

        strItemTag = BuildRoomTag(colTagValues.Item(CONST_TagPrefixAttributeName), strLocTagSeqNo, colTagValues.Item(CONST_TagSuffixAttributeName))

        UpdateRoomTag = True

        If Len(strLocTagSeqNo) > 0 Then
            'Check for uniqueness
            bUniqueStatus = CheckForRoomUnique(Datasource, Item, strItemTag, strDupMessage)

            If bUniqueStatus = False Then

                If m_isUIEnabled Then
                    intResponse = MsgBox(strDupMessage & " " & My.Resources.str5009, MsgBoxStyle.YesNo + MsgBoxStyle.SystemModal + MsgBoxStyle.Exclamation, My.Resources.str5001)
                Else
                    intResponse = MsgBoxResult.Yes
                End If

                If intResponse = MsgBoxResult.Yes Then
                    blnUnique = False

                    Do While blnUnique = False

                        'Get the next Available tag sequence number from Options Manager
                        strLocTagSeqNo = GetNextAvailTagSeqNo(Datasource, CONST_NameAttributeName, CONST_RoomNextSeqNoAttributeName)

                        strItemTag = BuildRoomTag(colTagValues.Item(CONST_TagPrefixAttributeName), strLocTagSeqNo, colTagValues.Item(CONST_TagSuffixAttributeName))

                        blnUnique = CheckForRoomUnique(Datasource, Item, strItemTag, strDupMessage)
                    Loop
                Else
                    varValue = Item.Attributes.Item(strPropertyName).Value
                    UpdateRoomTag = False
                End If

            End If

        End If

        'Update the values
        If UpdateRoomTag Then

            With Item
                If Len(strLocTagSeqNo) > 0 Then
                    .Attributes.Item(CONST_TagSequenceNoAttributeName).Value = strLocTagSeqNo
                Else
                    .Attributes.Item(CONST_TagSequenceNoAttributeName).Value = System.DBNull.Value
                End If

                If strLocTagSeqNo = "" Then
                    .ItemTag = System.DBNull.Value
                Else
                    If Len(colTagValues.Item(CONST_TagPrefixAttributeName)) > 0 Then
                        .TagPrefix = colTagValues.Item(CONST_TagPrefixAttributeName)
                    End If
                    If Len(colTagValues.Item(CONST_TagSuffixAttributeName)) > 0 Then
                        .TagSuffix = colTagValues.Item(CONST_TagSuffixAttributeName)
                    End If

                    If Len(strItemTag) > 0 Then
                        .ItemTag = strItemTag
                    End If
                End If

                If Not IsDBNull(varValue) Then
                    'trim off any spaces in the input value
                    varValue = Trim(varValue)
                End If
                .Commit()

            End With

        End If

        Exit Function

ErrorHandler:
        LogError("ItemTag - ItemTagFunc::UpdateRoomTag --> " & Err.Description)
        colTagValues = Nothing
        objFilter = Nothing

    End Function

    Private Function CheckForRoomUnique(ByRef Datasource As Llama.LMADataSource, ByRef Item As Llama.LMRoom, ByRef strItemTag As String, ByRef strMessage As String) As Boolean
        Dim bUniqueStatus As Boolean
        Dim objFilter As Llama.LMAFilter
        Dim strDupMessage As String

        On Error GoTo ErrorHandler
        CheckForRoomUnique = False
        bUniqueStatus = False
        strMessage = ""
        strDupMessage = ""

        If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuiltAndProjs Then
            strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Room " & My.Resources.str5011 'AsBuiltPlant or its projects
            If Len(Item.PartOfPlantItemID.ToString()) > 0 Then
                If Item.IsBulkItemIndex = CONST_TrueIndex Then
                    objFilter = GetFilter(CONST_RoomItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PartofPlantItem_SP_IDAttributeName, Item.PartOfPlantItemID, CONST_PlantItemTypeAttributeName, CStr(CONST_RoomPlantItemIndex), CONST_IsBulkItemAttributeName, CStr(CONST_FalseIndex))
                Else
                    objFilter = GetFilter(CONST_RoomItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PartofPlantItem_SP_IDAttributeName, Item.PartOfPlantItemID, CONST_PlantItemTypeAttributeName, CStr(CONST_RoomPlantItemIndex))
                End If
            Else
                If Item.IsBulkItemIndex = CONST_TrueIndex Then
                    objFilter = GetFilter(CONST_RoomItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_RoomPlantItemIndex), CONST_IsBulkItemAttributeName, CStr(CONST_FalseIndex))
                Else
                    objFilter = GetFilter(CONST_RoomItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_RoomPlantItemIndex))
                End If
            End If
            bUniqueStatus = CheckForGlobalUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex), strItemTag, "")
        Else
            'Check the local plant or project
            If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuilt Then
                strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Room " & My.Resources.str5010 'Active project
            Else
                strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Room " & My.Resources.str5007 'Active plant
            End If
            If Len(Item.PartOfPlantItemID.ToString()) > 0 Then
                objFilter = GetFilter(CONST_RoomItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PartofPlantItem_SP_IDAttributeName, Item.PartOfPlantItemID, CONST_ItemStatus, Const_ItemStatusValue)
            Else
                objFilter = GetFilter(CONST_RoomItemName, CONST_ItemTagAttributeName, strItemTag, CONST_ItemStatus, Const_ItemStatusValue)
            End If
            bUniqueStatus = CheckForUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex))

            'Do we need to check the AsBuilt too?
            If bUniqueStatus = True Then
                If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuilt Then
                    strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Room " & My.Resources.str5008 'found in AsBuilt Plant
                    If Len(Item.PartOfPlantItemID.ToString()) > 0 Then
                        If Item.IsBulkItemIndex = CONST_TrueIndex Then
                            objFilter = GetFilter(CONST_RoomItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PartofPlantItem_SP_IDAttributeName, Item.PartOfPlantItemID, CONST_PlantItemTypeAttributeName, CStr(CONST_RoomPlantItemIndex), CONST_IsBulkItemAttributeName, CStr(CONST_FalseIndex))
                        Else
                            objFilter = GetFilter(CONST_RoomItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PartofPlantItem_SP_IDAttributeName, Item.PartOfPlantItemID, CONST_PlantItemTypeAttributeName, CStr(CONST_RoomPlantItemIndex))
                        End If
                    Else
                        If Item.IsBulkItemIndex = CONST_TrueIndex Then
                            objFilter = GetFilter(CONST_RoomItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_RoomPlantItemIndex), CONST_IsBulkItemAttributeName, CStr(CONST_FalseIndex))
                        Else
                            objFilter = GetFilter(CONST_RoomItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_RoomPlantItemIndex))
                        End If
                    End If
                    bUniqueStatus = CheckForAsBuiltUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex), strItemTag, "")
                End If
            End If
        End If

        CheckForRoomUnique = bUniqueStatus
        If CheckForRoomUnique = False Then
            strMessage = strDupMessage
        End If
        Exit Function

ErrorHandler:
        LogError("ItemTag - ItemTagFunc::CheckForRoomUnique --> " & Err.Description)
        objFilter = Nothing

    End Function

    Private Function BuildRoomTag(ByVal strPrefix As String, ByVal strTagSeqNo As String, ByVal strSuffix As String) As String

        BuildRoomTag = strPrefix & "-" & strTagSeqNo & strSuffix

    End Function

    Private Function UpdateDuctRunTag(ByRef Datasource As Llama.LMADataSource, ByRef Item As Llama.LMDuctRun, ByRef varValue As Object, ByRef strPropertyName As String) As Boolean

        Dim colTagValues As Collection
        Dim objPlantGroups As Llama.LMPlantGroups
        Dim objPlantGroup As Llama.LMPlantGroup
        Dim objUnit As Llama.LMUnit
        Dim objFilter As Llama.LMAFilter
        Dim objDuctRun As Llama.LMDuctRun

        Dim strItemTag As String
        Dim strLocTagSeqNo As String
        Dim strDupMessage As String
        Dim strUnitName As String
        Dim intResponse As Short
        Dim blnUnique As Boolean
        Dim bUniqueStatus As Boolean

        On Error GoTo ErrorHandler

        strUnitName = ""
        strDupMessage = ""

        Item.Attributes.BuildAttributesOnDemand = True

        'Get the UnitCode Value
        objDuctRun = Datasource.GetDuctRun((Item.Id))
        objDuctRun.Attributes.BuildAttributesOnDemand = True

        On Error Resume Next
        objPlantGroup = objDuctRun.PlantGroupObject
        On Error GoTo ErrorHandler
        If Not objPlantGroup Is Nothing Then
            objPlantGroup.Attributes.BuildAttributesOnDemand = True

            If objPlantGroup.PlantGroupTypeIndex = 65 Then
                objUnit = Datasource.GetUnit((objPlantGroup.Id))
                objUnit.Attributes.BuildAttributesOnDemand = True
                'Unit code may not be set if the Units are created by retrieving from EF
                If Not IsDBNull(objUnit.UnitCode) Then
                    strUnitName = objUnit.UnitCode
                End If
            End If
        End If

        'Use the collections unique index feature to avoid a bunch of if statements to
        'determine what property was passed in
        colTagValues = New Collection

        With colTagValues

            .Add(Trim(VariantToString(varValue)), strPropertyName)

            Select Case strPropertyName

                Case CONST_TagSuffixAttributeName
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)
                    .Add(Trim(VariantToString((Item.OperFluidCode))), CONST_OperFluidCodeAttributeName)

                Case CONST_TagSequenceNoAttributeName
                    .Add(Trim(VariantToString((Item.TagSuffix))), CONST_TagSuffixAttributeName)
                    .Add(Trim(VariantToString((Item.OperFluidCode))), CONST_OperFluidCodeAttributeName)

                Case CONST_OperFluidCodeAttributeName
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)
                    .Add(Trim(VariantToString((Item.TagSuffix))), CONST_TagSuffixAttributeName)

                Case Else
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)
                    .Add(Trim(VariantToString((Item.TagSuffix))), CONST_TagSuffixAttributeName)
                    .Add(Trim(VariantToString((Item.OperFluidCode))), CONST_OperFluidCodeAttributeName)

            End Select

        End With

        strLocTagSeqNo = colTagValues.Item(CONST_TagSequenceNoAttributeName)

        If Len(strLocTagSeqNo) = 0 And strPropertyName <> CONST_TagSequenceNoAttributeName Then
            'Get the next tagseqno from Options Manager
            strLocTagSeqNo = GetNextAvailTagSeqNo(Datasource, CONST_NameAttributeName, CONST_DuctRunNextSeqNoAttributeName)
        End If

        strItemTag = BuildDuctRunTag(strUnitName, strLocTagSeqNo, colTagValues.Item(CONST_TagSuffixAttributeName), colTagValues.Item(CONST_OperFluidCodeAttributeName))

        UpdateDuctRunTag = True

        If Len(strLocTagSeqNo) > 0 Then
            objFilter = Nothing
            strDupMessage = My.Resources.str5000
            If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuiltAndProjs Then
                strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Duct Run " & My.Resources.str5011
                If Item.IsBulkItemIndex = CONST_TrueIndex Then
                    objFilter = GetFilter(CONST_DuctRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_DuctRunPlantItemIndex), CONST_IsBulkItemAttributeName, CStr(CONST_FalseIndex))
                Else
                    objFilter = GetFilter(CONST_DuctRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_DuctRunPlantItemIndex))
                End If
                bUniqueStatus = CheckForGlobalUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex), strItemTag, "")
            Else
                If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuilt Then
                    'strDupMessage = "Duplicate Tag " & strItemTag & " found in active project."
                    strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Duct Run " & My.Resources.str5010
                Else
                    'strDupMessage = "Duplicate Tag " & strItemTag & " found in active plant."
                    strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Duct Run " & My.Resources.str5007
                End If

                'Check Active Plant first
                objFilter = GetFilter(CONST_DuctRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_ItemStatus, Const_ItemStatusValue)
                bUniqueStatus = CheckForUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex))

                If bUniqueStatus = True Then
                    'strDupMessage = "Duplicate Tag " & strItemTag & " found in AsBuilt plant."
                    strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Duct Run " & My.Resources.str5008
                    If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuilt Then
                        objFilter = Nothing
                        If Item.IsBulkItemIndex = CONST_TrueIndex Then
                            objFilter = GetFilter(CONST_DuctRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_DuctRunPlantItemIndex), CONST_IsBulkItemAttributeName, CStr(CONST_FalseIndex))
                        Else
                            objFilter = GetFilter(CONST_DuctRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_DuctRunPlantItemIndex))
                        End If
                        bUniqueStatus = CheckForAsBuiltUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex), strItemTag, "")
                    End If
                End If
            End If

            If bUniqueStatus = False Then
                If m_isUIEnabled Then
                    intResponse = MsgBox(strDupMessage & " " & My.Resources.str5009, MsgBoxStyle.YesNo + MsgBoxStyle.SystemModal + MsgBoxStyle.Exclamation, My.Resources.str5001)
                Else
                    intResponse = MsgBoxResult.No
                End If

                If intResponse = MsgBoxResult.No Then

                    UpdateDuctRunTag = True

                    With Item
                        If Len(strLocTagSeqNo) > 0 Then
                            .Attributes.Item(CONST_TagSequenceNoAttributeName).Value = strLocTagSeqNo
                        End If

                        If Len(colTagValues.Item(CONST_TagSuffixAttributeName)) > 0 Then
                            .Attributes.Item(CONST_TagSuffixAttributeName).Value = colTagValues.Item(CONST_TagSuffixAttributeName)
                        End If
                        If Len(strItemTag) > 0 Then
                            .ItemTag = strItemTag
                        End If
                        .Commit()
                    End With

                    Exit Function

                ElseIf intResponse = MsgBoxResult.Yes Then

                    blnUnique = False

                    Do While blnUnique = False

                        'Get the next tagseqno from Options Manager
                        strLocTagSeqNo = GetNextAvailTagSeqNo(Datasource, CONST_NameAttributeName, CONST_DuctRunNextSeqNoAttributeName)

                        strItemTag = BuildDuctRunTag(strUnitName, strLocTagSeqNo, colTagValues.Item(CONST_TagSuffixAttributeName), colTagValues.Item(CONST_OperFluidCodeAttributeName))

                        objFilter = Nothing
                        If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuiltAndProjs Then
                            strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Duct Run " & My.Resources.str5011
                            If Item.IsBulkItemIndex = CONST_TrueIndex Then
                                'if the Ductrun is bulk its okay to have a duplicate on another bulk
                                'but not on a not-bulk.
                                objFilter = GetFilter(CONST_DuctRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_DuctRunPlantItemIndex), CONST_IsBulkItemAttributeName, CStr(CONST_FalseIndex))
                            Else
                                objFilter = GetFilter(CONST_DuctRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_DuctRunPlantItemIndex))
                            End If
                            blnUnique = CheckForGlobalUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex), strItemTag, "")
                        Else
                            objFilter = GetFilter(CONST_DuctRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_ItemStatus, Const_ItemStatusValue)
                            blnUnique = CheckForUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex))

                            If blnUnique = True Then
                                'Its unique in the project - check in the Asbuilt
                                If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuilt Then
                                    objFilter = Nothing
                                    If Item.IsBulkItemIndex = CONST_TrueIndex Then
                                        objFilter = GetFilter(CONST_DuctRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_DuctRunPlantItemIndex), CONST_IsBulkItemAttributeName, CStr(CONST_FalseIndex))
                                    Else
                                        objFilter = GetFilter(CONST_DuctRunItemName, CONST_ItemTagAttributeName, strItemTag, CONST_PlantItemTypeAttributeName, CStr(CONST_DuctRunPlantItemIndex))
                                    End If
                                    blnUnique = CheckForAsBuiltUnique(Datasource, objFilter, CStr(Item.Id), (Item.IsBulkItemIndex), strItemTag, "")
                                End If
                            End If
                        End If
                    Loop

                Else
                    UpdateDuctRunTag = False
                End If

            End If

        End If

        'Update the item values
        If UpdateDuctRunTag Then

            With Item

                If Len(strLocTagSeqNo) > 0 Then
                    .Attributes.Item(CONST_TagSequenceNoAttributeName).Value = strLocTagSeqNo
                Else
                    .Attributes.Item(CONST_TagSequenceNoAttributeName).Value = System.DBNull.Value
                End If

                If strLocTagSeqNo = "" Then
                    If Len(colTagValues.Item(CONST_OperFluidCodeAttributeName)) = 0 Then
                        .ItemTag = System.DBNull.Value
                    Else
                        If Len(colTagValues.Item(CONST_TagSuffixAttributeName)) > 0 Then
                            .Attributes.Item(CONST_TagSuffixAttributeName).Value = colTagValues.Item(CONST_TagSuffixAttributeName)
                        End If
                        If Len(strItemTag) > 0 Then
                            .ItemTag = strItemTag
                        End If
                    End If

                    .ItemTag = System.DBNull.Value
                Else
                    If Len(colTagValues.Item(CONST_TagSuffixAttributeName)) > 0 Then
                        .Attributes.Item(CONST_TagSuffixAttributeName).Value = colTagValues.Item(CONST_TagSuffixAttributeName)
                    End If
                    If Len(strItemTag) > 0 Then
                        .ItemTag = strItemTag
                    End If
                End If

                If Not IsDBNull(varValue) Then
                    'trim off any spaces in the input value
                    varValue = Trim(varValue)
                End If

                .Commit()

            End With

        End If

        Exit Function

ErrorHandler:
        LogError("ItemTag - ItemTagFunc::UpdateDuctRunTag --> " & Err.Description)
        objDuctRun = Nothing
        objPlantGroups = Nothing
        objUnit = Nothing
        colTagValues = Nothing
        objFilter = Nothing

    End Function

    Private Function IsRoomAttached(ByRef Item As Llama.LMNozzle, ByRef Datasource As Llama.LMADataSource) As Boolean

        Dim objLMPlantItem As Llama.LMPlantItem
        Dim vPartOfID As Object
        Dim strID As String
        Dim res As Boolean

        On Error GoTo errHandler

        res = False
        objLMPlantItem = Datasource.GetPlantItem((Item.Id))
        objLMPlantItem.Attributes.BuildAttributesOnDemand = True

        vPartOfID = objLMPlantItem.PartOfPlantItemID

        While Not IsDBNull(vPartOfID)
            objLMPlantItem = objLMPlantItem.PartOfPlantItemObject
            objLMPlantItem.Attributes.BuildAttributesOnDemand = True

            vPartOfID = objLMPlantItem.PartOfPlantItemID
        End While

        If objLMPlantItem.PlantItemTypeIndex = 21 Then
            strID = CStr(objLMPlantItem.Attributes(CONST_SP_RoomIDAttributeName).Value)
            res = True
        End If

        IsRoomAttached = res
        Exit Function

errHandler:
        res = False
        IsRoomAttached = res

    End Function

    Private Function UpdateNozzleTagForRoom(ByRef Datasource As Llama.LMADataSource, ByRef Item As Llama.LMNozzle, ByRef varValue As Object, ByRef strPropertyName As String) As Boolean

        Dim objLMPlantItem As Llama.LMPlantItem
        Dim objChildPlantItem As Llama.LMPlantItem
        Dim objLMRoom As Llama.LMRoom
        Dim objRoomComp As Llama.LMRoomComponents
        Dim objLMNozzle As Llama.LMNozzle
        Dim vPartOfID As Object
        Dim strRoomID As String
        Dim ItemTag_Renamed As String
        Dim locTagSeqNo As String
        Dim nBulkItemIndex As Integer
        Dim Response As Short
        Dim colTagValues As Collection
        Dim strDupMessage As String
        Dim strRoomTag As String
        Dim strLocTagSeqNo As String
        Dim strItemTag As String
        Dim intResponse As Short
        Dim DoUntilUni As Boolean
        Dim objlmaitems As Llama.LMAItems
        Dim objFilter As Llama.LMAFilter
        Dim strLocTagPrefix As String
        Dim strLocTagSuffix As String
        Dim bAttachToRoom As Boolean

        On Error GoTo errHandler

        UpdateNozzleTagForRoom = True


        'In following cases do not increment TagSeqNo
        '1.Placement
        '2.Drag&Drop when nozzle don't have item tag
        If strPropertyName = CONST_ItemTagAttributeName And Len("" & varValue) = 0 Then Exit Function


        colTagValues = New Collection
        strDupMessage = ""
        strRoomTag = ""

        Item.Attributes.BuildAttributesOnDemand = True

        With colTagValues

            .Add(Trim(VariantToString(varValue)), strPropertyName)

            Select Case strPropertyName

                Case CONST_TagSuffixAttributeName
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagPrefixAttributeName).Value))), CONST_TagPrefixAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)

                Case CONST_TagPrefixAttributeName
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSuffixAttributeName).Value))), CONST_TagSuffixAttributeName)

                Case CONST_TagSequenceNoAttributeName
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagPrefixAttributeName).Value))), CONST_TagPrefixAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSuffixAttributeName).Value))), CONST_TagSuffixAttributeName)

                Case CONST_ItemTagAttributeName
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagPrefixAttributeName).Value))), CONST_TagPrefixAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value))), CONST_TagSequenceNoAttributeName)
                    .Add(Trim(VariantToString((Item.Attributes.Item(CONST_TagSuffixAttributeName).Value))), CONST_TagSuffixAttributeName)

                Case Else

            End Select

        End With

        strLocTagSeqNo = colTagValues.Item(CONST_TagSequenceNoAttributeName)
        strItemTag = BuildNozzleTag(colTagValues.Item(CONST_TagPrefixAttributeName), strLocTagSeqNo, colTagValues.Item(CONST_TagSuffixAttributeName))

        If strPropertyName = CONST_TagSequenceNoAttributeName And strLocTagSeqNo = "" Then GoTo ExitDueTOTagSeq
        If strLocTagSeqNo = "" Then
            If colTagValues.Item(CONST_TagSuffixAttributeName) <> "" Then
                DoUntilUni = True
            Else
                If colTagValues.Item(CONST_TagPrefixAttributeName) <> "" Then
                    DoUntilUni = True
                End If
            End If
        End If

        strLocTagPrefix = colTagValues.Item(CONST_TagPrefixAttributeName)
        If strPropertyName = CONST_TagPrefixAttributeName And strLocTagPrefix <> "" Then
            If strLocTagSeqNo = "" Then
                strLocTagSeqNo = CStr(1)
            End If
        End If
        strLocTagSuffix = colTagValues.Item(CONST_TagSuffixAttributeName)
        If strPropertyName = CONST_TagSuffixAttributeName And strLocTagSuffix <> "" Then
            If strLocTagSeqNo = "" Then
                strLocTagSeqNo = CStr(1)
            End If
        End If

        strItemTag = BuildNozzleTag(colTagValues.Item(CONST_TagPrefixAttributeName), strLocTagSeqNo, colTagValues.Item(CONST_TagSuffixAttributeName))

        nBulkItemIndex = Item.IsBulkItemIndex

        'get the RoomID from the Nozzle. Here Nozzle can be on Room Component.
        'so navigate upto the Room.
        If IsDBNull(Item.RoomID) Then

            objLMPlantItem = Datasource.GetPlantItem((Item.Id))
            objLMPlantItem.Attributes.BuildAttributesOnDemand = True

            vPartOfID = objLMPlantItem.PartOfPlantItemID

            While Not IsDBNull(vPartOfID)
                objLMPlantItem = objLMPlantItem.PartOfPlantItemObject
                objLMPlantItem.Attributes.BuildAttributesOnDemand = True

                vPartOfID = objLMPlantItem.PartOfPlantItemID
            End While
            If objLMPlantItem.PlantItemTypeIndex = 21 Then
                strRoomID = CStr(objLMPlantItem.Attributes(CONST_SP_RoomIDAttributeName).Value)
            Else
                strRoomID = CStr(objLMPlantItem.Id)
            End If
        Else
            strRoomID = CStr(Item.RoomID)
        End If

        'get the Room
        objLMRoom = Datasource.GetRoom(strRoomID)
        objLMRoom.Attributes.BuildAttributesOnDemand = True
        If Not IsDBNull(objLMRoom.Attributes("ItemTag").Value) Then
            strRoomTag = objLMRoom.Attributes("ItemTag").Value
        End If

MainLabel:
        If intResponse = MsgBoxResult.Yes Then
            If m_NozzleSeqNo = 0 Then m_NozzleSeqNo = 1
            'reset the starting Nozzle Tag Seq No to 1 if the Room is different
            If Len(m_PrevNozzleEqID) = 0 Then
                m_PrevNozzleEqID = objLMRoom.Id
            ElseIf m_PrevNozzleEqID <> objLMRoom.Id Then
                m_PrevNozzleEqID = objLMRoom.Id
                'reset the next TagSeqNo to 1 if the Room is different.
                m_NozzleSeqNo = 1
            End If

            strLocTagSeqNo = CStr(m_NozzleSeqNo + 1)
            m_NozzleSeqNo = CInt(strLocTagSeqNo)
            strItemTag = BuildNozzleTag(colTagValues.Item(CONST_TagPrefixAttributeName), strLocTagSeqNo, colTagValues.Item(CONST_TagSuffixAttributeName))
        End If

        strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Nozzle on Room " & strRoomTag & " " & My.Resources.str5007
        m_objLMNozzles = New Llama.LMNozzles

        For Each objLMNozzle In objLMRoom.Nozzles
            m_objLMNozzles.Add(objLMNozzle.AsLMAItem)
        Next objLMNozzle

        ' The nozzles on the children
        For Each objLMPlantItem In objLMRoom.ChildPlantItemPlantItems
            Call GetChildNozzles(objLMPlantItem)
        Next objLMPlantItem

        For Each objLMNozzle In m_objLMNozzles
            objLMNozzle.Attributes.BuildAttributesOnDemand = True

            If objLMNozzle.ItemStatusIndex <> 4 Then
                If Not IsDBNull(objLMNozzle.ItemTag) Then
                    If objLMNozzle.ItemTag = strItemTag Then
                        If objLMNozzle.IsBulkItemIndex = CONST_TrueIndex And nBulkItemIndex = CONST_TrueIndex Or objLMNozzle.Id = Item.Id Then
                            UpdateNozzleTagForRoom = True
                        Else
                            UpdateNozzleTagForRoom = False
                            'Exit For
                            GoTo ExitDueToDupInActivePlant
                        End If
                    End If
                End If
            End If
        Next objLMNozzle

        If UpdateNozzleTagForRoom = True Then
            If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuilt Then
                strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Nozzle on Room " & strRoomTag & " " & My.Resources.str5008
                objFilter = Nothing
                objFilter = GetFilter(CONST_NozzleItemName, CONST_ItemTagAttributeName, strItemTag)
                UpdateNozzleTagForRoom = CheckForAsBuiltUnique(Datasource, objFilter, CStr(Item.Id), nBulkItemIndex, strItemTag, strRoomID)
            End If
        End If

        If UpdateNozzleTagForRoom = True Then
            If m_DuplicateTagCheckScope = eDuplicateTagCheckScope.ActiveProjAgainstAsBuiltAndProjs Then
                strDupMessage = My.Resources.str5006 & " '" & strItemTag & "' for Nozzle on Room " & strRoomTag & " " & My.Resources.str5011
                objFilter = Nothing
                objFilter = GetFilter(CONST_NozzleItemName, CONST_ItemTagAttributeName, strItemTag)
                UpdateNozzleTagForRoom = CheckForGlobalUnique(Datasource, objFilter, CStr(Item.Id), nBulkItemIndex, strItemTag, strRoomID)
            End If
        End If

ExitDueToDupInActivePlant:

        If UpdateNozzleTagForRoom = True Then
ExitDueTOTagSeq:
            Item.Attributes.Item(CONST_ItemTagAttributeName).Value = strItemTag
            If Len(strLocTagSeqNo) > 0 Then
                Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value = strLocTagSeqNo
            Else
                Item.Attributes.Item(CONST_TagSequenceNoAttributeName).Value = System.DBNull.Value
            End If

            If Not IsDBNull(varValue) Then
                'trim off any spaces in the input value
                varValue = Trim(varValue)
            End If

            If strPropertyName = CONST_TagPrefixAttributeName Then
                Item.Attributes.Item(CONST_TagPrefixAttributeName).Value = varValue
                If Len(colTagValues.Item(CONST_TagSuffixAttributeName)) > 0 Then
                    Item.Attributes.Item(CONST_TagSuffixAttributeName).Value = colTagValues.Item(CONST_TagSuffixAttributeName)
                End If
            ElseIf strPropertyName = CONST_TagSuffixAttributeName Then
                Item.Attributes.Item(CONST_TagSuffixAttributeName).Value = varValue
                If Len(colTagValues.Item(CONST_TagPrefixAttributeName)) > 0 Then
                    Item.Attributes.Item(CONST_TagPrefixAttributeName).Value = colTagValues.Item(CONST_TagPrefixAttributeName)
                End If
            Else
                If Len(colTagValues.Item(CONST_TagPrefixAttributeName)) > 0 Then
                    Item.Attributes.Item(CONST_TagPrefixAttributeName).Value = colTagValues.Item(CONST_TagPrefixAttributeName)
                End If
                If Len(colTagValues.Item(CONST_TagSuffixAttributeName)) > 0 Then
                    Item.Attributes.Item(CONST_TagSuffixAttributeName).Value = colTagValues.Item(CONST_TagSuffixAttributeName)
                End If
            End If
            Item.Commit()
            UpdateNozzleTagForRoom = True
            Exit Function
        Else
            If m_isUIEnabled Then
                If intResponse = MsgBoxResult.Yes Then
                    intResponse = MsgBoxResult.Yes
                Else
                    If DoUntilUni = True Then
                        intResponse = MsgBoxResult.Yes
                    Else
                        If strDupMessage = "" Then
                            intResponse = MsgBox(My.Resources.str5000, MsgBoxStyle.YesNo + MsgBoxStyle.SystemModal + MsgBoxStyle.Exclamation, My.Resources.str5001)
                        Else
                            intResponse = MsgBox(strDupMessage & " " & My.Resources.str5009, MsgBoxStyle.YesNo + MsgBoxStyle.SystemModal + MsgBoxStyle.Exclamation, My.Resources.str5001)
                        End If
                    End If
                End If
            Else
                intResponse = MsgBoxResult.Yes
            End If
            If intResponse = MsgBoxResult.Yes Then
                UpdateNozzleTagForRoom = True
                GoTo MainLabel
            Else
                UpdateNozzleTagForRoom = False
            End If
        End If
        'cleanup
        objLMPlantItem = Nothing
        objChildPlantItem = Nothing
        objLMRoom = Nothing
        objRoomComp = Nothing
        objLMNozzle = Nothing
        vPartOfID = System.DBNull.Value
        objFilter = Nothing
        Exit Function

errHandler:
        'cleanup
        objLMPlantItem = Nothing
        objChildPlantItem = Nothing
        objLMRoom = Nothing
        objRoomComp = Nothing
        objLMNozzle = Nothing
        vPartOfID = System.DBNull.Value
        objFilter = Nothing
    End Function
End Class
