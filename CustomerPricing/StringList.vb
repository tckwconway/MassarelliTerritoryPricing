
'==========================================================
' * 
' * Currently contained classes:
' * StringList -- Encapsulates an ArrayList to deal specificaly with strings and string[]
' * 
' * 
' * 
' 
'==========================================================

Imports System.Collections

'Helper classes meant to speed up development 
Namespace JTG
    ''' <summary>
    ''' Summary description for Class1.
    ''' </summary>
    ''' 

    Public Class StringList
        Private m_alMain As ArrayList
        'constructor
        Public Sub New()
            m_alMain = New ArrayList()
        End Sub
        'destructor
        Protected Overrides Sub Finalize()
            Try
                m_alMain.Clear()
            Finally
                MyBase.Finalize()
            End Try
        End Sub
        'add an item to the end of the list
        'returns the number of items in the list
        Public Function Add(s As String) As Integer
            m_alMain.Add(s)
            Return m_alMain.Count
        End Function
        Public Function AddRange(sl As StringList) As StringList
            m_alMain.AddRange(sl.m_alMain)
            Return Nothing
        End Function
        'insert a string into the list only if the same string hasn't been added yet
        Public Function AddNoDuplicate(s As String) As Integer
            If m_alMain.IndexOf(s) = -1 Then
                m_alMain.Add(s)
            End If
            Return m_alMain.Count
        End Function
        'insert an item at the desired position
        'will throw an error if index is out of range
        'returns the number of items in the list
        Public Function Insert(index As Integer, s As String) As Integer
            m_alMain.Insert(index, s)
            Return m_alMain.Count
        End Function
        'remove an item from the list
        'returns the number of items in the list
        Public Function Remove(s As [String]) As Integer
            m_alMain.Remove(s)
            Return m_alMain.Count
        End Function
        Public Function Replace(sFind As String, sReplace As String) As Integer
            Dim index As Integer = m_alMain.IndexOf(sFind)
            If index > -1 Then
                m_alMain.RemoveAt(index)
                m_alMain.Insert(index, sReplace)
            End If
            Return m_alMain.Count
        End Function
        'remove all items from the list
        Public Sub Clear()
            m_alMain.Clear()
        End Sub
        'indexer
        Default Public Property Item(index As Integer) As [String]
            Get
                Return DirectCast(m_alMain(index), String)
            End Get
            Set(value As [String])
                If index >= m_alMain.Count Then
                    m_alMain.Add(value)
                Else
                    m_alMain(index) = value
                End If
            End Set
        End Property
        'ToString() override
        'Return the contents of the list with the items seperated by a new line
        Public Overrides Function ToString() As String
            Dim sRHS As String = ""
            Dim index As Integer = 0
            While index < m_alMain.Count
                sRHS += DirectCast(m_alMain(System.Math.Max(System.Threading.Interlocked.Increment(index), index - 1)), String) & Convert.ToString(vbLf)
            End While
            Return sRHS
        End Function
        'return the contents of the list seperated by the given seperator string
        Public Overloads Function ToString(sSeperator As String) As [String]
            Dim sRHS As String = ""
            Dim index As Integer = 0
            While index < m_alMain.Count
                sRHS += DirectCast(m_alMain(System.Math.Max(System.Threading.Interlocked.Increment(index), index - 1)), String)
                If index < m_alMain.Count Then
                    sRHS += sSeperator
                End If
            End While
            Return sRHS
        End Function
        'Sort the string in the array
        Public Sub Sort()
            m_alMain.Sort()
        End Sub
        'Properties
        'Count
        'returns the number of items in the list
        Public ReadOnly Property Count() As Integer
            Get
                Return m_alMain.Count
            End Get
        End Property
        ''' <summary>
        ''' Finds the index of a given string.
        ''' </summary>
        ''' <param name="sFind">String to find the index of.</param>
        ''' <returns>The index of sFind or -1 if it doesn't exist.</returns>
        Public Function IndexOf(sFind As String) As Integer
            Return m_alMain.IndexOf(sFind)
        End Function
        'conversion operators
        'convert a StringList to a string[]
        Public Shared Widening Operator CType(sl As StringList) As String()
            Dim sLHS As String() = New String(sl.m_alMain.Count - 1) {}
            Dim index As Integer = 0
            While index < sl.m_alMain.Count
                sLHS(index) = DirectCast(sl.m_alMain(index), String)
                index += 1
            End While
            Return sLHS
        End Operator
        'convert a string[] to a StringList
        Public Shared Widening Operator CType(sa As String()) As StringList
            Dim sl As New StringList()
            For index As Integer = 0 To sa.Length - 1
                sl.Add(sa(index))
            Next
            Return sl
        End Operator
        'convert a string[] to a StringList
        Public Shared Widening Operator CType(sa As Object()) As StringList
            Dim sl As New StringList()
            For index As Integer = 0 To sa.Length - 1
                sl.Add(sa(index).ToString())
            Next
            Return sl
        End Operator
    End Class
End Namespace

'=======================================================
'Service provided by Telerik (www.telerik.com)
'Conversion powered by NRefactory.
'Twitter: @telerik
'Facebook: facebook.com/telerik
'=======================================================
