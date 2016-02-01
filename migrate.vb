Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE90
Imports System.Diagnostics

Public Module ProjectUtilities

    Private Class ProjectGuids
        Public Const vsWindowsCSharp As String = "{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}"
        Public Const vsWindowsVBNET As String = "{F184B08F-C81C-45F6-A57F-5ABD9991F28F}"
        Public Const vsWindowsVisualCPP As String = "{8BC9CEB8-8B4A-11D0-8D11-00A0C91BC942}"
        Public Const vsWebApplication As String = "{349C5851-65DF-11DA-9384-00065B846F21}"
        Public Const vsWebSite As String = "{E24C65DC-7377-472B-9ABA-BC803B73C61A}"
        Public Const vsDistributedSystem As String = "{F135691A-BF7E-435D-8960-F99683D2D49C}"
        Public Const vsWCF As String = "{3D9AD99F-2412-4246-B90B-4EAA41C64699}"
        Public Const vsWPF As String = "{60DC8134-EBA5-43B8-BCC9-BB4BC16C2548}"
        Public Const vsVisualDatabaseTools As String = "{C252FEB5-A946-4202-B1D4-9916A0590387}"
        Public Const vsDatabase As String = "{A9ACE9BB-CECE-4E62-9AA4-C7E7C5BD2124}"
        Public Const vsDatabaseOther As String = "{4F174C21-8C12-11D0-8340-0000F80270F8}"
        Public Const vsTest As String = "{3AC096D0-A1C2-E12C-1390-A8335801FDAB}"
        Public Const vsLegacy2003SmartDeviceCSharp As String = "{20D4826A-C6FA-45DB-90F4-C717570B9F32}"
        Public Const vsLegacy2003SmartDeviceVBNET As String = "{CB4CE8C6-1BDB-4DC7-A4D3-65A1999772F8}"
        Public Const vsSmartDeviceCSharp As String = "{4D628B5B-2FBC-4AA6-8C16-197242AEB884}"
        Public Const vsSmartDeviceVBNET As String = "{68B1623D-7FB9-47D8-8664-7ECEA3297D4F}"
        Public Const vsWorkflowCSharp As String = "{14822709-B5A1-4724-98CA-57A101D1B079}"
        Public Const vsWorkflowVBNET As String = "{D59BE175-2ED0-4C54-BE3D-CDAA9F3214C8}"
        Public Const vsDeploymentMergeModule As String = "{06A35CCD-C46D-44D5-987B-CF40FF872267}"
        Public Const vsDeploymentCab As String = "{3EA9E505-35AC-4774-B492-AD1749C4943A}"
        Public Const vsDeploymentSetup As String = "{978C614F-708E-4E1A-B201-565925725DBA}"
        Public Const vsDeploymentSmartDeviceCab As String = "{AB322303-2255-48EF-A496-5904EB18DA55}"
        Public Const vsVSTA As String = "{A860303F-1F3F-4691-B57E-529FC101A107}"
        Public Const vsVSTO As String = "{BAA0C2D2-18E2-41B9-852F-F413020CAA33}"
        Public Const vsSharePointWorkflow As String = "{F8810EC1-6754-47FC-A15F-DFABD2E3FA90}"
    End Class

    '' Defines the valid target framework values.
    Enum TargetFramework
        Fx40 = 262144
        Fx35 = 196613
        Fx30 = 196608
        Fx20 = 131072
    End Enum

    '' Change the target framework for all projects in the current solution.
    Sub ChangeTargetFrameworkForAllProjects()
        Dim project As EnvDTE.Project
        Dim clientProfile As Boolean = False

        Write("--------- CHANGING TARGET .NET FRAMEWORK VERSION -------------")
        Try
            If Not DTE.Solution.IsOpen Then
                Write("There is no solution open.")
            Else              
                Dim targetFrameworkInput As String = InputBox("Enter the target framework version (Fx40, Fx35, Fx30, Fx20):", "Target Framework", "Fx40")
                Dim targetFramework As TargetFramework = [Enum].Parse(GetType(TargetFramework), targetFrameworkInput)

                If targetFramework = ProjectUtilities.TargetFramework.Fx35 Or targetFramework = ProjectUtilities.TargetFramework.Fx40 Then
                    Dim result As MsgBoxResult = MsgBox("The .NET Framework version chosen supports a Client Profile. Would you like to use that profile?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, "Target Framework Profile")
                    If result = MsgBoxResult.Yes Then
                        clientProfile = True
                    End If
                End If

                For Each project In DTE.Solution.Projects
                    If project.Kind <> Constants.vsProjectKindSolutionItems And project.Kind <> Constants.vsProjectKindMisc Then
                        ChangeTargetFramework(project, targetFramework, clientProfile)
                    Else
                        For Each projectItem In project.ProjectItems
                            If Not (projectItem.SubProject Is Nothing) Then
                                ChangeTargetFramework(projectItem.SubProject, targetFramework, clientProfile)
                            End If
                        Next

                    End If
                Next
            End If
        Catch ex As System.Exception
            Write(ex.Message)
        End Try
    End Sub

    '' Change the target framework for a project.
    Function ChangeTargetFramework(ByVal project As EnvDTE.Project, ByVal targetFramework As TargetFramework, ByVal clientProfile As Boolean) As Boolean
        Dim changed As Boolean = True

        If project.Kind = Constants.vsProjectKindSolutionItems Or project.Kind = Constants.vsProjectKindMisc Then
            For Each projectItem In project.ProjectItems
                If Not (projectItem.SubProject Is Nothing) Then
                    ChangeTargetFramework(projectItem.SubProject, targetFramework, clientProfile)
                End If
            Next
        Else
            Try
                If IsLegalProjectType(project) Then
                    SetTargetFramework(project, targetFramework, clientProfile)
                Else
                    Write("Skipping project: " + project.Name + " (" + project.Kind + ")")
                End If
            Catch ex As Exception
                Write(ex.Message)
                changed = False
            End Try
        End If

        Return changed
    End Function

    '' Determines if the project is a project that actually supports changing the target framework.
    Function IsLegalProjectType(ByVal proejct As EnvDTE.Project) As Boolean
        Dim legalProjectType As Boolean = True

        Select Case proejct.Kind
            Case ProjectGuids.vsDatabase
                legalProjectType = False
            Case ProjectGuids.vsDatabaseOther
                legalProjectType = False
            Case ProjectGuids.vsDeploymentCab
                legalProjectType = False
            Case ProjectGuids.vsDeploymentMergeModule
                legalProjectType = False
            Case ProjectGuids.vsDeploymentSetup
                legalProjectType = False
            Case ProjectGuids.vsDeploymentSmartDeviceCab
                legalProjectType = False
            Case ProjectGuids.vsDistributedSystem
                legalProjectType = False
            Case ProjectGuids.vsLegacy2003SmartDeviceCSharp
                legalProjectType = False
            Case ProjectGuids.vsLegacy2003SmartDeviceVBNET
                legalProjectType = False
            Case ProjectGuids.vsSharePointWorkflow
                legalProjectType = False
            Case ProjectGuids.vsSmartDeviceCSharp
                legalProjectType = True
            Case ProjectGuids.vsSmartDeviceVBNET
                legalProjectType = True
            Case ProjectGuids.vsTest
                legalProjectType = False
            Case ProjectGuids.vsVisualDatabaseTools
                legalProjectType = False
            Case ProjectGuids.vsVSTA
                legalProjectType = True
            Case ProjectGuids.vsVSTO
                legalProjectType = True
            Case ProjectGuids.vsWCF
                legalProjectType = True
            Case ProjectGuids.vsWebApplication
                legalProjectType = True
            Case ProjectGuids.vsWebSite
                legalProjectType = True
            Case ProjectGuids.vsWindowsCSharp
                legalProjectType = True
            Case ProjectGuids.vsWindowsVBNET
                legalProjectType = True
            Case ProjectGuids.vsWindowsVisualCPP
                legalProjectType = True
            Case ProjectGuids.vsWorkflowCSharp
                legalProjectType = False
            Case ProjectGuids.vsWorkflowVBNET
                legalProjectType = False
            Case ProjectGuids.vsWPF
                legalProjectType = True
            Case Else
                legalProjectType = False
        End Select
        Return legalProjectType
    End Function

    '' Sets the target framework for the project to the specified framework.
    Sub SetTargetFramework(ByVal project As EnvDTE.Project, ByVal targetFramework As TargetFramework, ByVal clientProfile As Boolean)
        Dim currentTargetFramework As TargetFramework = CType(project.Properties.Item("TargetFramework").Value, TargetFramework)
        Dim targetMoniker As String = GetTargetFrameworkMoniker(targetFramework, clientProfile)
        Dim currentMoniker As String = project.Properties.Item("TargetFrameworkMoniker").Value

        If currentMoniker <> targetMoniker Then
            Write("Changing project: " + project.Name + " from " + currentMoniker + " to " + targetMoniker + ".")
            project.Properties.Item("TargetFrameworkMoniker").Value = targetMoniker
            project.Properties.Item("TargetFramework").Value = targetFramework
        Else
            Write("Skipping project: " + project.Name + ", already at the correct target framework.")
        End If
    End Sub

    Function GetTargetFrameworkMoniker(ByVal targetFramework As TargetFramework, ByVal clientProfile As Boolean) As String
        Dim moniker As String = ".NETFramework,Version=v"
        Select Case targetFramework
            Case ProjectUtilities.TargetFramework.Fx20
                moniker += "2.0"

            Case ProjectUtilities.TargetFramework.Fx30
                moniker += "3.0"

            Case ProjectUtilities.TargetFramework.Fx35
                moniker += "3.5"

            Case ProjectUtilities.TargetFramework.Fx40
                moniker += "4.0"

        End Select

        If clientProfile Then
            moniker += ",Profile=Client"
        End If

        Return moniker
    End Function

    '' Writes a message to the output window
    Sub Write(ByVal s As String)
        Dim out As OutputWindowPane = GetOutputWindowPane("Change Target Framework", True)
        out.OutputString(s)
        out.OutputString(vbCrLf)
    End Sub

    '' Gets an instance of the output window
    Function GetOutputWindowPane(ByVal Name As String, Optional ByVal show As Boolean = True) As OutputWindowPane
        Dim win As Window = DTE.Windows.Item(EnvDTE.Constants.vsWindowKindOutput)
        If show Then win.Visible = True
        Dim ow As OutputWindow = win.Object
        Dim owpane As OutputWindowPane
        Try
            owpane = ow.OutputWindowPanes.Item(Name)
        Catch e As System.Exception
            owpane = ow.OutputWindowPanes.Add(Name)
        End Try
        owpane.Activate()
        Return owpane
    End Function

End Module
