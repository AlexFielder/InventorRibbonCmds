Imports System
Imports System.Environment

Imports NewRibbonCmds.StandardAddInServer

'a class for reading and writing XML style application settings
'sample usage:
'  Create the object -
'      Dim a As New AppSettings(AppSettings.Config.PrivateFile)
'  Retrieve a setting -
'      Dim myString As String = a.GetSetting("SettingName")
'  Save a setting -
'      a.SaveSetting("SettingName", value)
'

Public Class AppSettings
    Private _configFileName As String 'local var to hold the config file name
    Private _configFileType As Config 'local var to hold the config file type (private or shared)

    'config file options
    Public Enum Config
        SharedFile  'all users use the same config file
        PrivateFile 'each user has their own config file
    End Enum

    'constructor
    Public Sub New(ByVal ConfigFileType As Config)
        _configFileType = ConfigFileType 'remember this setting

        InitializeConfigFile() 'setup the filename and location
    End Sub

    'initialize the apps config file, create it if it doesn't exist
    Private Sub InitializeConfigFile()
        Dim sb As New Text.StringBuilder
        '
        'finish building the file name
        sb.Append(CurrentDirectory)
        sb.Append("\")
        sb.Append(ThisApplication.ProductName & "_config")
        sb.Append(".xml")
        '
        _configFileName = sb.ToString 'completed config filename
        '
        'if the file doesn't exist, create a blank xml
        If Not IO.File.Exists(_configFileName) Then
            Dim fn As New IO.StreamWriter(IO.File.Open(_configFileName, IO.FileMode.Create))
            fn.WriteLine("<?xml version=""1.0"" encoding=""utf-8""?>")
            fn.WriteLine("<configuration>")
            fn.WriteLine("  <appSettings>")
            fn.WriteLine("    <!--   User application and configured property settings go here.-->")
            fn.WriteLine("    <!--   Example: <add key=""settingName"" value=""settingValue""/> -->")
            fn.WriteLine("  </appSettings>")
            fn.WriteLine("</configuration>")
            fn.Close() 'all done
        End If
        ''
    End Sub

    'get an application setting by key value
    Public Function GetSetting(ByVal key As String) As String
        'xml document object
        Dim xd As New Xml.XmlDocument

        'load the xml file
        xd.Load(_configFileName)

        'query for a value
        Dim Node As Xml.XmlNode = xd.DocumentElement.SelectSingleNode( _
                                  "/configuration/appSettings/add[@key=""" & key & """]")

        'return the value or nothing if it doesn't exist
        If Not Node Is Nothing Then
            Return Node.Attributes.GetNamedItem("value").Value
        Else
            Return Nothing
        End If
    End Function

    'save an application setting, takes a key and a value
    Public Sub SaveSetting(ByVal key As String, ByVal value As String)
        'xml document object
        Dim xd As New Xml.XmlDocument

        'load the xml file
        xd.Load(_configFileName)

        'get the value
        Dim Node As Xml.XmlElement = CType(xd.DocumentElement.SelectSingleNode( _
                                           "/configuration/appSettings/add[@key=""" & _
                                           key & """]"), Xml.XmlElement)
        If Not Node Is Nothing Then
            'key found, set the value
            Node.Attributes.GetNamedItem("value").Value = value
        Else
            'key not found, create it
            Node = xd.CreateElement("add")
            Node.SetAttribute("key", key)
            Node.SetAttribute("value", value)

            'look for the appsettings node
            Dim Root As Xml.XmlNode = xd.DocumentElement.SelectSingleNode("/configuration/appSettings")

            'add the new child node (this key)
            If Not Root Is Nothing Then
                Root.AppendChild(Node)
            Else
                Try
                    'appsettings node didn't exist, add it before adding the new child
                    Root = xd.DocumentElement.SelectSingleNode("/configuration")
                    Root.AppendChild(xd.CreateElement("appSettings"))
                    Root = xd.DocumentElement.SelectSingleNode("/configuration/appSettings")
                    Root.AppendChild(Node)
                Catch ex As Exception
                    'failed adding node, throw an error
                    Throw New Exception("Could not set value", ex)
                End Try
            End If
        End If

        'finally, save the new version of the config file
        xd.Save(_configFileName)
    End Sub
End Class