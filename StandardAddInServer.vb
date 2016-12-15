Imports Inventor
Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Namespace SheetMetalExtents
    <ProgIdAttribute("SheetMetalExtents.StandardAddInServer"), _
    GuidAttribute("2c264a08-501d-476f-8414-8a52331a8046")> _
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer

        ' Inventor application object.
        Private m_inventorApplication As Inventor.Application
        Private WithEvents m_applicationEvents As Inventor.ApplicationEvents
        Private WithEvents m_styleEvents As Inventor.StyleEvents

        Private m_widthName As String = "SheetMetalWidth"
        Private m_lengthName As String = "SheetMetalLength"
        Private m_styleName As String = "SheetMetalStyle"

        Private m_processing As Boolean = False


#Region "ApplicationAddInServer Members"

        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate
            m_inventorApplication = addInSiteObject.Application
            m_applicationEvents = m_inventorApplication.ApplicationEvents
            m_styleEvents = m_inventorApplication.StyleEvents
        End Sub

        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate
            ' Release global objects.
            m_styleEvents = Nothing
            m_applicationEvents = Nothing
            m_inventorApplication = Nothing
        End Sub

        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation
            Get
                Return Nothing
            End Get
        End Property

        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand
            ' Not used.
        End Sub
#End Region

        ' Handle the OnOpenDocument event to make sure the newly opened document is up to date.
        Private Sub m_applicationEvents_OnOpenDocument(ByVal DocumentObject As Inventor._Document, ByVal FullDocumentName As String, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_applicationEvents.OnOpenDocument
            Try
                If BeforeOrAfter = EventTimingEnum.kAfter Then
                    ' Check that the document is a sheet metal document.
                    If DocumentObject.DocumentSubType.DocumentSubTypeID = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
                        Dim invSheetMetalDoc As Inventor.PartDocument = CType(DocumentObject, Inventor.PartDocument)

                        ' Check to see if the model has changed since the last update.
                        Dim needsAttribute As Boolean = True
                        Dim needsUpdate As Boolean = False
                        If invSheetMetalDoc.ComponentDefinition.AttributeSets.NameIsUsed("ekinsSheetMetalExtents") Then
                            Dim attribset As Inventor.AttributeSet
                            attribset = invSheetMetalDoc.ComponentDefinition.AttributeSets.Item("ekinsSheetMetalExtents")

                            ' Check that this attribute exists.  It always should, if the attribute set exists
                            ' but there was one case where it didn't.
                            If attribset.NameIsUsed("ModelGeometryVersion") Then
                                ' Compare the saved model geometry version with the current version.
                                If invSheetMetalDoc.ComponentDefinition.ModelGeometryVersion <> CType(attribset.Item("ModelGeometryVersion").Value, String) Then
                                    ' The versions are different so it need to be updated.
                                    ' Set the flags to cause an update.
                                    needsUpdate = True

                                    ' Update the saved version value with the current value.
                                    m_processing = True
                                    attribset.Item("ModelGeometryVersion").Value = invSheetMetalDoc.ComponentDefinition.ModelGeometryVersion
                                    m_processing = False
                                End If

                                needsAttribute = False
                            Else
                                ' There's a problem with the attribute set so delete it to allow it
                                ' to be recreated below.
                                attribset.Delete()
                            End If
                        End If

                        If needsAttribute Then
                            ' The attribute doesn't exist so create it to track the B-Rep changes.
                            Dim attribSet As Inventor.AttributeSet
                            m_processing = True
                            attribSet = invSheetMetalDoc.ComponentDefinition.AttributeSets.Add("ekinsSheetMetalExtents")
                            attribSet.Add("ModelGeometryVersion", ValueTypeEnum.kStringType, invSheetMetalDoc.ComponentDefinition.ModelGeometryVersion)
                            m_processing = False
                            needsUpdate = True
                        End If

                        ' Update the properties, if needed.
                        If needsUpdate Then
                            UpdatePropertyValues(invSheetMetalDoc)
                        End If
                    End If
                End If
            Catch ex As Exception
                ' Do nothing.
            End Try
        End Sub


        ' Handle the OnDocumentChange event to monitor any changes that might affect the extents of the flat pattern.
        Private Sub m_applicationEvents_OnDocumentChange(ByVal DocumentObject As Inventor._Document, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal ReasonsForChange As Inventor.CommandTypesEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_applicationEvents.OnDocumentChange
            If Not m_processing Then
                If BeforeOrAfter = EventTimingEnum.kAfter Then
                    Try
                        ' Check to see if it's a sheet metal document.
                        If DocumentObject.DocumentSubType.DocumentSubTypeID = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
                            Dim invSheetMetalDoc As Inventor.PartDocument = CType(DocumentObject, Inventor.PartDocument)

                            ' Check to see if it's an "kUpdateWithReferencesCmdType" or "kShapeEditCmdType" type of command.
                            ' This turned out not to be reliable because a change and returns from a sketch that causes a 
                            ' compute is being labeled as a query cmd and not a shape edit.  Adding support for query to catch it.
                            If (ReasonsForChange And CommandTypesEnum.kUpdateWithReferencesCmdType) = CommandTypesEnum.kUpdateWithReferencesCmdType Or _
                                (ReasonsForChange And CommandTypesEnum.kShapeEditCmdType) = CommandTypesEnum.kShapeEditCmdType Or _
                                (ReasonsForChange And CommandTypesEnum.kQueryOnlyCmdType) = CommandTypesEnum.kQueryOnlyCmdType Then

                                Dim needsUpdate As Boolean = False

                                ' Check to see if the model has changed since the last update.
                                If invSheetMetalDoc.ComponentDefinition.AttributeSets.NameIsUsed("ekinsSheetMetalExtents") Then
                                    Dim attribset As Inventor.AttributeSet
                                    attribset = invSheetMetalDoc.ComponentDefinition.AttributeSets.Item("ekinsSheetMetalExtents")

                                    If invSheetMetalDoc.ComponentDefinition.ModelGeometryVersion <> CType(attribset.Item("ModelGeometryVersion").Value, String) Then
                                        needsUpdate = True
                                        m_processing = True
                                        attribset.Item("ModelGeometryVersion").Value = invSheetMetalDoc.ComponentDefinition.ModelGeometryVersion
                                        m_processing = False
                                    End If
                                Else
                                    ' Create the attribute to track the B-Rep changes.
                                    Dim attribSet As Inventor.AttributeSet
                                    m_processing = True
                                    attribSet = invSheetMetalDoc.ComponentDefinition.AttributeSets.Add("ekinsSheetMetalExtents")
                                    attribSet.Add("ModelGeometryVersion", ValueTypeEnum.kStringType, invSheetMetalDoc.ComponentDefinition.ModelGeometryVersion)
                                    m_processing = False
                                    needsUpdate = True
                                End If

                                If Not needsUpdate Then
                                    Try
                                        ' Check to see if the flat pattern was deleted.
                                        If CType(Context.Value("InternalName"), String) = "CompositeChange" Then
                                            Dim commandList As Object
                                            commandList = Context.Value("InternalNamesList")
                                            Dim tempArray As String()
                                            tempArray = CType(commandList, String())

                                            If tempArray(0) = "Delete FlatPattern" Then
                                                needsUpdate = True
                                            End If
                                        End If
                                    Catch ex As Exception
                                        ' Do nothing
                                    End Try
                                End If

                                If needsUpdate Then
                                    UpdatePropertyValues(invSheetMetalDoc)
                                End If
                            ElseIf (ReasonsForChange And CommandTypesEnum.kNonShapeEditCmdType) = CommandTypesEnum.kNonShapeEditCmdType Then
                                If CType(Context.Value("InternalName"), String) = "SetSheetMetalDefaults" Then
                                    UpdatePropertyValues(invSheetMetalDoc)
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        ' Do nothing.
                    End Try
                End If
            End If
        End Sub

        Private Sub UpdatePropertyValues(ByVal PartDocument As Inventor.PartDocument)
            Dim transaction As Inventor.Transaction = Nothing
            Dim madeChange As Boolean = False
            Try
                m_processing = True

                ' Start a transaction so all of these changes are packaged in a single undo.
                transaction = m_inventorApplication.TransactionManager.StartTransaction(CType(PartDocument, Inventor._Document), "Flat Pattern Property Update")

                Dim sheetMetalCompDef As SheetMetalComponentDefinition = CType(PartDocument.ComponentDefinition, Inventor.SheetMetalComponentDefinition)

                ' Update the properties.
                Dim invCustomPropSet As PropertySet
                invCustomPropSet = PartDocument.PropertySets.Item("Inventor User Defined Properties")

                ' Get one property and two parameters, if they exist or create them if they don't.
                Dim invStyleProperty As Inventor.Property = Nothing
                Dim invLengthParameter As Inventor.UserParameter = Nothing
                Dim invWidthParameter As Inventor.UserParameter = Nothing

                Dim sheetMetalStyleName As String = sheetMetalCompDef.ActiveSheetMetalStyle.Name
                Dim errorUpdating As Boolean = False
                Try
                    invStyleProperty = invCustomPropSet.Item(m_styleName)

                    ' Update the style if needed.
                    If sheetMetalStyleName <> CType(invStyleProperty.Value, String) Then
                        invStyleProperty.Value = sheetMetalStyleName
                        madeChange = True
                    End If
                Catch ex As Exception
                    errorUpdating = True
                End Try

                ' Unable to update.  Assume it doesn't exist and create it.
                If errorUpdating Then
                    invStyleProperty = invCustomPropSet.Add(sheetMetalStyleName, m_styleName)
                    madeChange = True
                End If

                ' Check to see if any of these parameters exist as reference parameters
                ' and convert them to user parameters to migrate files that used the previous version of
                ' this program that created reference parameters.
                Dim tempRefParam As Inventor.ReferenceParameter = Nothing
                Try
                    tempRefParam = sheetMetalCompDef.Parameters.ReferenceParameters(m_widthName)
                Catch ex As Exception
                    ' Do Nothing
                End Try

                If Not tempRefParam Is Nothing Then
                    tempRefParam.ConvertToUserParameter()
                End If

                Try
                    tempRefParam = sheetMetalCompDef.Parameters.ReferenceParameters(m_lengthName)
                Catch ex As Exception
                    ' Do Nothing
                End Try

                If Not tempRefParam Is Nothing Then
                    tempRefParam.ConvertToUserParameter()
                End If

                ' Create the parameters, if needed.
                Dim paramExists As Boolean = True
                Try
                    invWidthParameter = sheetMetalCompDef.Parameters.UserParameters.Item(m_widthName)
                Catch ex As Exception
                    paramExists = False
                End Try

                If Not paramExists Then
                    invWidthParameter = sheetMetalCompDef.Parameters.UserParameters.AddByValue(m_widthName, 0, UnitsTypeEnum.kDefaultDisplayLengthUnits)
                    invWidthParameter.ExposedAsProperty = True
                    madeChange = True
                End If

                paramExists = True
                Try
                    invLengthParameter = sheetMetalCompDef.Parameters.UserParameters.Item(m_lengthName)
                Catch ex As Exception
                    paramExists = False
                End Try

                If Not paramExists Then
                    invLengthParameter = sheetMetalCompDef.Parameters.UserParameters.AddByValue(m_lengthName, 0, UnitsTypeEnum.kDefaultDisplayLengthUnits)
                    invLengthParameter.ExposedAsProperty = True
                    madeChange = True
                End If

                ' Check to see if the flat exists.
                If sheetMetalCompDef.HasFlatPattern Then
                    ' Get the flat pattern.
                    Dim invFlatPattern As FlatPattern
                    invFlatPattern = sheetMetalCompDef.FlatPattern

                    ' Update the parameter values if they've changed.
                    If Math.Abs(invFlatPattern.Length - CType(invLengthParameter.Value, Double)) > 0.0000001 Then
                        invLengthParameter.Value = invFlatPattern.Length
                        madeChange = True
                    End If

                    If Math.Abs(invFlatPattern.Width - CType(invWidthParameter.Value, Double)) > 0.0000001 Then
                        invWidthParameter.Value = invFlatPattern.Width
                        madeChange = True
                    End If
                Else
                    invLengthParameter.Value = 0
                    invWidthParameter.Value = 0
                    madeChange = True
                End If

                If madeChange Then
                    transaction.End()

                    Try
                        transaction.MergeWithPrevious = True
                    Catch ex As Exception
                    End Try
                Else
                    transaction.Abort()
                End If
            Catch ex As Exception
                If Not transaction Is Nothing Then
                    transaction.Abort()
                End If

                If Not m_inventorApplication.SilentOperation Then
                    MsgBox("Unexpected error while updating the property values with the sheet metal extents.  The results should not be used.", MsgBoxStyle.Information And MsgBoxStyle.OkOnly)
                End If
            Finally
                m_processing = False
            End Try
        End Sub


        'Private Sub SetReferenceParamValue(ByVal ReferenceParam As Inventor.ReferenceParameter, ByVal NewValue As Double)
        '    Try
        '        Dim userParam As Inventor.UserParameter = ReferenceParam.ConvertToUserParameter
        '        userParam.Value = NewValue
        '        userParam.ConvertToReferenceParameter()
        '    Catch ex As Exception
        '        ' Do nothing
        '    End Try
        'End Sub


        Private Sub m_styleEvents_OnActivateStyle(DocumentObject As _Document, Style As Object, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum) Handles m_styleEvents.OnActivateStyle
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                If TypeOf Style Is SheetMetalStyle Then
                    If DocumentObject.DocumentSubType.DocumentSubTypeID = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
                        Dim invSheetMetalDoc As Inventor.PartDocument = CType(DocumentObject, Inventor.PartDocument)
                        UpdatePropertyValues(invSheetMetalDoc)
                    End If
                End If
            End If
        End Sub
    End Class
End Namespace

