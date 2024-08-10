Imports System.IO
Imports System.Text.RegularExpressions

Namespace DracLabs
    Public Class IniFile
        Private m_sections As List(Of IniSection)

        Public Sub New()
            m_sections = New List(Of IniSection)()
        End Sub

        ''' <summary>
        ''' Loads the data from the ini file into the IniFile object.
        ''' </summary>
        ''' <param name="sFileName">The path to the ini file.</param>
        ''' <param name="bMerge">If true, merges the data with existing sections.</param>
        Public Sub Load(ByVal sFileName As String, Optional ByVal bMerge As Boolean = False)
            If Not bMerge Then
                RemoveAllSections()
            End If
            Dim tempsection As IniSection = Nothing
            Dim oReader As New StreamReader(sFileName)
            Dim regexcomment As New Regex("^([\s]*#.*)", (RegexOptions.Singleline Or RegexOptions.IgnoreCase))
            Dim iLineNumber As Integer = 0
            Dim regexsection As New Regex("^[\s]*\[[\s]*([^\[\s].*[^\s\]])[\s]*\][\s]*$", (RegexOptions.Singleline Or RegexOptions.IgnoreCase))
            Dim regexkey As New Regex("^\s*([^=\s]*)[^=]*=(.*)", (RegexOptions.Singleline Or RegexOptions.IgnoreCase))
            While Not oReader.EndOfStream
                Dim line As String = oReader.ReadLine()
                iLineNumber += 1
                If line <> String.Empty Then
                    If line.Length > 10000 Then
                        line = Left(line, 10000)
                    End If
                    Dim m As Match = Nothing
                    If regexcomment.Match(line).Success Then
                        m = regexcomment.Match(line)
                    ElseIf regexsection.Match(line).Success Then
                        m = regexsection.Match(line)
                        tempsection = AddSection(m.Groups(1).Value)
                    ElseIf tempsection IsNot Nothing AndAlso regexkey.Match(line).Success Then
                        m = regexkey.Match(line)
                        tempsection.AddKey(m.Groups(1).Value).Value = m.Groups(2).Value
                    ElseIf tempsection IsNot Nothing Then
                        tempsection.AddKey(line)
                    End If
                End If
            End While
            oReader.Close()
        End Sub

        ''' <summary>
        ''' Saves the data back to the specified ini file.
        ''' </summary>
        ''' <param name="sFileName">The path to the ini file.</param>
        Public Sub Save(ByVal sFileName As String)
            Dim oWriter As New StreamWriter(sFileName, False)
            For Each s As IniSection In m_sections
                oWriter.WriteLine(String.Format("[{0}]", s.Name))
                For Each k As IniSection.IniKey In s.Keys
                    If k.Value <> String.Empty Then
                        oWriter.WriteLine(String.Format("{0}={1}", k.Name, k.Value))
                    Else
                        oWriter.WriteLine(String.Format("{0}", k.Name))
                    End If
                Next
            Next
            oWriter.Close()
        End Sub

        ''' <summary>
        ''' Gets all the sections.
        ''' </summary>
        Public ReadOnly Property Sections() As System.Collections.ICollection
            Get
                Return m_sections
            End Get
        End Property

        ''' <summary>
        ''' Adds a section to the IniFile object.
        ''' </summary>
        ''' <param name="sSection">The name of the section.</param>
        ''' <returns>The added or existing IniSection object.</returns>
        Public Function AddSection(ByVal sSection As String) As IniSection
            Dim s As IniSection = GetSection(sSection)
            If s Is Nothing Then
                s = New IniSection(Me, sSection)
                m_sections.Add(s)
            End If
            Return s
        End Function

        ''' <summary>
        ''' Removes a section by its name.
        ''' </summary>
        ''' <param name="sSection">The name of the section.</param>
        ''' <returns>True if the section was removed, otherwise false.</returns>
        Public Function RemoveSection(ByVal sSection As String) As Boolean
            Dim s As IniSection = GetSection(sSection)
            If s IsNot Nothing Then
                m_sections.Remove(s)
                Return True
            End If
            Return False
        End Function

        ''' <summary>
        ''' Removes all existing sections.
        ''' </summary>
        ''' <returns>True if all sections were removed, otherwise false.</returns>
        Public Function RemoveAllSections() As Boolean
            m_sections.Clear()
            Return (m_sections.Count = 0)
        End Function

        ''' <summary>
        ''' Returns an IniSection by name.
        ''' </summary>
        ''' <param name="sSection">The name of the section.</param>
        ''' <returns>The IniSection object if found, otherwise null.</returns>
        Public Function GetSection(ByVal sSection As String) As IniSection
            For Each s As IniSection In m_sections
                If s.Name.Equals(sSection, StringComparison.InvariantCultureIgnoreCase) Then
                    Return s
                End If
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' Returns a key value in a certain section.
        ''' </summary>
        ''' <param name="sSection">The name of the section.</param>
        ''' <param name="sKey">The name of the key.</param>
        ''' <returns>The value of the key if found, otherwise an empty string.</returns>
        Public Function GetKeyValue(ByVal sSection As String, ByVal sKey As String) As String
            Dim s As IniSection = GetSection(sSection)
            If s IsNot Nothing Then
                Dim k As IniSection.IniKey = s.GetKey(sKey)
                If k IsNot Nothing Then
                    Return k.Value
                End If
            End If
            Return String.Empty
        End Function

        ''' <summary>
        ''' Sets a key value pair in a certain section.
        ''' </summary>
        ''' <param name="sSection">The name of the section.</param>
        ''' <param name="sKey">The name of the key.</param>
        ''' <param name="sValue">The value of the key.</param>
        ''' <returns>True if the key value pair was set, otherwise false.</returns>
        Public Function SetKeyValue(ByVal sSection As String, ByVal sKey As String, ByVal sValue As String) As Boolean
            Dim s As IniSection = AddSection(sSection)
            If s IsNot Nothing Then
                Dim k As IniSection.IniKey = s.AddKey(sKey)
                If k IsNot Nothing Then
                    k.Value = sValue
                    Return True
                End If
            End If
            Return False
        End Function

        ''' <summary>
        ''' Renames an existing section.
        ''' </summary>
        ''' <param name="sSection">The current name of the section.</param>
        ''' <param name="sNewSection">The new name of the section.</param>
        ''' <returns>True if the section was renamed, otherwise false.</returns>
        Public Function RenameSection(ByVal sSection As String, ByVal sNewSection As String) As Boolean
            Dim s As IniSection = GetSection(sSection)
            If s IsNot Nothing Then
                Return s.SetName(sNewSection)
            End If
            Return False
        End Function

        ''' <summary>
        ''' Renames an existing key.
        ''' </summary>
        ''' <param name="sSection">The name of the section.</param>
        ''' <param name="sKey">The current name of the key.</param>
        ''' <param name="sNewKey">The new name of the key.</param>
        ''' <returns>True if the key was renamed, otherwise false.</returns>
        Public Function RenameKey(ByVal sSection As String, ByVal sKey As String, ByVal sNewKey As String) As Boolean
            Dim s As IniSection = GetSection(sSection)
            If s IsNot Nothing Then
                Dim k As IniSection.IniKey = s.GetKey(sKey)
                If k IsNot Nothing Then
                    Return k.SetName(sNewKey)
                End If
            End If
            Return False
        End Function

        Public Class IniSection
            Private m_pIniFile As IniFile
            Private m_sSection As String
            Private m_keys As List(Of IniKey)

            Protected Friend Sub New(ByVal parent As IniFile, ByVal sSection As String)
                m_pIniFile = parent
                m_sSection = sSection
                m_keys = New List(Of IniKey)()
            End Sub

            ''' <summary>
            ''' Returns all the keys in a section.
            ''' </summary>
            Public ReadOnly Property Keys() As System.Collections.ICollection
                Get
                    Return m_keys
                End Get
            End Property

            ''' <summary>
            ''' Returns the section name.
            ''' </summary>
            Public ReadOnly Property Name() As String
                Get
                    Return m_sSection
                End Get
            End Property

            ''' <summary>
            ''' Adds a key to the IniSection object.
            ''' </summary>
            ''' <param name="sKey">The name of the key.</param>
                        ''' <returns>The added or existing IniKey object.</returns>
            Public Function AddKey(ByVal sKey As String) As IniKey
                sKey = sKey.Trim()
                Dim k As IniKey = GetKey(sKey)
                If k Is Nothing Then
                    k = New IniKey(Me, sKey)
                    m_keys.Add(k)
                End If
                Return k
            End Function

            ''' <summary>
            ''' Removes a single key by its name.
            ''' </summary>
            ''' <param name="sKey">The name of the key.</param>
            ''' <returns>True if the key was removed, otherwise false.</returns>
            Public Function RemoveKey(ByVal sKey As String) As Boolean
                Dim k As IniKey = GetKey(sKey)
                If k IsNot Nothing Then
                    m_keys.Remove(k)
                    Return True
                End If
                Return False
            End Function

            ''' <summary>
            ''' Removes all the keys in the section.
            ''' </summary>
            ''' <returns>True if all keys were removed, otherwise false.</returns>
            Public Function RemoveAllKeys() As Boolean
                m_keys.Clear()
                Return (m_keys.Count = 0)
            End Function

            ''' <summary>
            ''' Returns an IniKey object by name.
            ''' </summary>
            ''' <param name="sKey">The name of the key.</param>
            ''' <returns>The IniKey object if found, otherwise null.</returns>
            Public Function GetKey(ByVal sKey As String) As IniKey
                For Each k As IniKey In m_keys
                    If k.Name.Equals(sKey, StringComparison.InvariantCultureIgnoreCase) Then
                        Return k
                    End If
                Next
                Return Nothing
            End Function

            ''' <summary>
            ''' Sets the section name.
            ''' </summary>
            ''' <param name="sSection">The new name of the section.</param>
            ''' <returns>True if the section name was set, otherwise false.</returns>
            Public Function SetName(ByVal sSection As String) As Boolean
                If m_pIniFile.GetSection(sSection) Is Nothing Then
                    m_sSection = sSection
                    Return True
                End If
                Return False
            End Function

            Public Class IniKey
                Private m_sKey As String
                Private m_sValue As String
                Private m_section As IniSection

                Protected Friend Sub New(ByVal parent As IniSection, ByVal sKey As String)
                    m_section = parent
                    m_sKey = sKey
                End Sub

                ''' <summary>
                ''' Returns the name of the key.
                ''' </summary>
                Public ReadOnly Property Name() As String
                    Get
                        Return m_sKey
                    End Get
                End Property

                ''' <summary>
                ''' Gets or sets the value of the key.
                ''' </summary>
                Public Property Value() As String
                    Get
                        Return m_sValue
                    End Get
                    Set(ByVal value As String)
                        m_sValue = value
                    End Set
                End Property

                ''' <summary>
                ''' Sets the key name.
                ''' </summary>
                ''' <param name="sKey">The new name of the key.</param>
                ''' <returns>True if the key name was set, otherwise false.</returns>
                Public Function SetName(ByVal sKey As String) As Boolean
                    If m_section.GetKey(sKey) Is Nothing Then
                        m_sKey = sKey
                        Return True
                    End If
                    Return False
                End Function
            End Class
        End Class
    End Class
End Namespace
