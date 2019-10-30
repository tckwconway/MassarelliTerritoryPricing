Imports System.Data
Imports System.Collections
Imports System.Drawing
Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.Runtime.InteropServices

Namespace JTG
    ''' <summary>
    ''' Summary description for ColumnComboBox.
    ''' </summary>
    Public Class ColumnComboBox
        Inherits ComboBox
        Private m_slSuggestions As New StringList()
        'A handy class used with the suggestion features.
        Private m_kcLastKey As Keys = Keys.Space
        'Last key pressed.
        Private m_iaColWidths As Integer() = New Integer(0) {}
        'Used for quick access to the column widths.
        Public ColumnSpacing As UInteger = 4
        'Minimum spacing between columns. Don't go crazy with this...
        Private m_Cols As New CCBColumnCollection()
        'A class used for managing the columns.
        Private m_dtData As DataTable = Nothing
        'Main DataTable and DataView that contain the information to be shown in the ColumnComboBox.
        Private m_dvView As DataView = Nothing

        'private bool m_bShowHeadings = true; //Was going to do something with this but ran out of time.
        Private m_iViewColumn As Integer = 0
        'Which of the columns will be shown in the text box.
        Private m_bInitItems As Boolean = True
        'Flags used to determine when the things need to be initialized.
        Private m_bInitDisplay As Boolean = True
        Private m_bInitSuggestionList As Boolean = True

        Private m_bTextChangedInternal As Boolean = False
        'Used when the text is being changed by another member of the class.
        Public DropDownOnSuggestion As Boolean = True
        Public Suggest As Boolean = True
        'Suggesting can be turned on or off. No need for the whole property write out.
        Private m_iSelectedIndex As Integer = -1
        'Used for storing the selected index without depending on the base.
        Private components As System.ComponentModel.Container = Nothing

        Public Sub New()

            If components Is Nothing Then
                components = Nothing
            End If
            Data = New DataTable()
            'Make sure the DataTable is not blank

            Init()
        End Sub
        Public Sub New(dtData As DataTable)
            Data = dtData
            Init()
        End Sub
        Private Sub Init()
            Me.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable

        End Sub

#Region "Component Designer generated code"
        ''' <summary>
        ''' Required method for Designer support - do not modify 
        ''' the contents of this method with the code editor.
        ''' </summary>
        Private Sub InitializeComponent()
            ' 
            ' ColumnComboBox
            ' 

        End Sub

#End Region

        Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
            Try
                If m_bInitSuggestionList Then
                    InitSuggestionList()
                End If
                MyBase.OnKeyDown(e)
                m_kcLastKey = e.KeyCode
            Catch ex As Exception
                Throw New Exception(ex.Message + vbCr & vbLf & "In ColumnComboBox.OnKeyDown(KeyEventArgs).")
            End Try
        End Sub

        Protected Overrides Sub OnTextChanged(e As EventArgs)
            'Doesn't call the base so no wiring up this event for you.
            Try
                'Run a few checks to make sure there should be any "suggesting" going on.
                If m_bTextChangedInternal Then
                    'If the text is being changed by another member of this class do nothing
                    m_bTextChangedInternal = False
                    'It will only be getting changed once internally, next time do something.
                    Return
                End If
                If Not Suggest Then
                    Return
                End If
                If SelectionStart < Me.Text.Length Then
                    Return
                End If
                Dim iOffset As Integer = 0
                If (m_kcLastKey = Keys.Back) OrElse (m_kcLastKey = Keys.Delete) Then
                    'Obviously we aren't going to find anything when they push Backspace or Delete
                    UpdateIndex()
                    Return
                End If
                If m_slSuggestions Is Nothing OrElse Me.Text.Length < 1 Then
                    Return
                End If

                'Put the current text into temp storage
                Dim sText As String
                sText = Me.Text
                Dim sOriginal As String = sText
                sText = sText.ToUpper()
                Dim iLength As Integer = sText.Length
                Dim sFound As String = Nothing
                Dim index As Integer = 0
                'see if what is currently in the text box matches anything in the string list
                For index = 0 To m_slSuggestions.Count - 1
                    Dim sTemp As String = m_slSuggestions(index).ToUpper()
                    If sTemp.Length >= sText.Length Then
                        If sTemp.IndexOf(sText, 0, sText.Length) > -1 Then
                            sFound = m_slSuggestions(index)
                            Exit For
                        End If
                    End If
                Next
                If sFound IsNot Nothing Then
                    m_bTextChangedInternal = True
                    If DropDownOnSuggestion AndAlso Not DroppedDown Then
                        m_bTextChangedInternal = True
                        Dim sTempText As String = Text
                        Me.DroppedDown = True
                        Text = sTempText
                        m_bTextChangedInternal = False
                    End If
                    If Me.Text <> sFound Then
                        Me.Text += sFound.Substring(iLength)
                        Me.SelectionStart = iLength + iOffset
                        Me.SelectionLength = Me.Text.Length - iLength + iOffset
                        m_iSelectedIndex = index
                        SelectedIndex = index
                        MyBase.OnSelectedIndexChanged(New EventArgs())
                    Else
                        UpdateIndex()
                        Me.SelectionStart = iLength
                        Me.SelectionLength = 0
                    End If
                Else
                    m_bTextChangedInternal = True
                    m_iSelectedIndex = -1
                    SelectedIndex = -1
                    Text = sOriginal
                    m_bTextChangedInternal = False
                    MyBase.OnSelectedIndexChanged(New EventArgs())
                    Me.SelectionStart = sOriginal.Length
                    Me.SelectionLength = 0
                End If
            Catch ex As Exception
                'Throw New Exception(ex.Message + vbCr & vbLf & "In ColumnComboBox.OnTextChanged(EventArgs).")
            End Try
        End Sub

        Protected Overrides Sub OnDropDown(e As EventArgs)
            Try
                'Initialize as required.
                If m_bInitItems Then
                    InitItems()
                End If
                If m_bInitDisplay Then
                    InitDisplay()
                End If
                MyBase.OnDropDown(e)
            Catch ex As Exception
                Throw New Exception(ex.Message + vbCr & vbLf & "In ColumnComboBox.OnDropDown(EventArgs).")
            End Try
        End Sub

        Protected Overrides Sub OnSelectedIndexChanged(e As EventArgs)
            Try
                'Keep track of this internally.

                m_iSelectedIndex = MyBase.SelectedIndex
                MyBase.OnSelectedIndexChanged(e)
            Catch ex As Exception
                'Throw New Exception(ex.Message + vbCr & vbLf & "In ColumnComboBox.OnSelectedIndexChanged(EventArgs).")
            End Try
        End Sub


        'This is where the magic happens that makes it appear dropped down with multiple columns
        Protected Overrides Sub OnDrawItem(e As DrawItemEventArgs)
            Try
                Dim iIndex As Integer = e.Index
                If iIndex > -1 Then
                    Dim iXPos As Integer = 0
                    Dim iYPos As Integer = 0

                    Dim dr As DataRow = m_dvView(iIndex).Row
                    e.DrawBackground()
                    For index As Integer = 0 To m_Cols.Count - 1
                        'Loop for drawing each column
                        If m_Cols(index).Display = False Then
                            Continue For
                        End If
                        e.Graphics.DrawString(dr(index).ToString(), Font, New SolidBrush(e.ForeColor), New RectangleF(iXPos, e.Bounds.Y, m_Cols(index).CalculatedWidth, ItemHeight))
                        iXPos += m_Cols(index).CalculatedWidth - 4
                    Next
                    iXPos = 0
                    iYPos += ItemHeight
                    e.DrawFocusRectangle()
                    MyBase.OnDrawItem(e)
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message + vbCr & vbLf & "In ColumnComboBox.OnDrawItem(DrawItemEventArgs).")
            End Try
        End Sub
        Private Sub InitItems()
            Try
                'Reset the Columns and the base.Items list
                m_Cols.Clear()
                For Each dc As DataColumn In m_dtData.Columns
                    m_Cols.Add(New CCBColumn(dc.Caption))
                Next
                'Set to the first or last column if an invlid ViewColumn is specified.
                If m_iViewColumn > m_Cols.Count - 1 Then
                    m_iViewColumn = m_Cols.Count - 1
                End If
                If m_iViewColumn < 0 Then
                    m_iViewColumn = 0
                End If

                'Set up the events for the columns
                For index As Integer = 0 To m_Cols.Count - 1
                    AddHandler m_Cols(index).OnColumnDisplayChanged, New ChangeColumnDisplayHandler(AddressOf ColumnComboBox_OnColumnDisplayChanged)
                Next

                MyBase.Items.Clear()
                'Put the stuff from the ViewColumn into the base so other base functionality will work
                For Each drv As DataRowView In m_dvView
                    Dim sTemp As String = drv(m_iViewColumn).ToString()
                    MyBase.Items.Add(sTemp)
                Next
                m_bInitItems = False
                'Set the flag to initialize the display before next drop down
                m_bInitDisplay = True
            Catch ex As Exception
                Throw New Exception(ex.Message + vbCr & vbLf & "In ColumnComboBox.InitItems().")
            End Try
        End Sub
        Private Sub InitDisplay()
            Try
                'Set the widths of the columns
                Dim m_iaColWidths As Integer() = New Integer(m_Cols.Count - 1) {}
                Dim size As New SizeF(10000, ItemHeight)
                'Here is a nice magic number for you but it should suffice.
                Dim g As Graphics = CreateGraphics()
                m_iaColWidths = New Integer(m_Cols.Count - 1) {}
                'Measure each column width and set the largest size needed for each column
                For Each drv As DataRowView In m_dvView
                    For index As Integer = 0 To m_Cols.Count - 1
                        Dim sTemp As String = drv(index).ToString()
                        Dim iTempWidth As Integer = CInt(g.MeasureString(sTemp, Font, size).Width)
                        If iTempWidth > m_iaColWidths(index) Then
                            m_iaColWidths(index) = iTempWidth
                        End If
                    Next
                Next
                DropDownWidth = 1
                For index As Integer = 0 To m_iaColWidths.Length - 1
                    If m_Cols(index).Width < 0 Then
                        'It will be < 0 if it hasn't been initialized.
                        m_Cols(index).CalculatedWidth = m_iaColWidths(index) + CInt(ColumnSpacing)
                    Else
                        m_Cols(index).CalculatedWidth = m_Cols(index).Width + CInt(ColumnSpacing)
                    End If
                    Dim a As Integer = 0
                    a += 1
                    If m_Cols(index).Display Then
                        DropDownWidth += m_Cols(index).CalculatedWidth
                    End If
                Next
                DropDownWidth += 16
                'Another nice magic number to represent the vertical scroll bar width
                m_bInitDisplay = False
            Catch ex As Exception
                Throw New Exception(ex.Message + vbCr & vbLf & "In ColumnComboBox.InitDisplay().")
            End Try
        End Sub

        'Put all the data from the ViewColumn into a StringList for quicker suggesting later.
        Private Sub InitSuggestionList()
            m_slSuggestions.Clear()
            For Each drv As DataRowView In m_dvView
                Dim sTemp As String = drv(m_iViewColumn).ToString()
                m_slSuggestions.Add(sTemp)
            Next
        End Sub

        'Sometimes you just have to command the ComboBox to update its SelectedIndex.
        'This function will do that based on the current text.
        Public Sub UpdateIndex()
            Try
                If m_bInitItems Then
                    InitItems()
                End If
                If m_bInitSuggestionList Then
                    InitSuggestionList()
                End If
                Dim sText As String = Text
                Dim index As Integer = 0
                For index = 0 To m_dvView.Count - 1
                    If m_dvView(index)(ViewColumn).ToString() = sText Then
                        If SelectedIndex <> index Then
                            m_bTextChangedInternal = True
                            m_iSelectedIndex = index
                            SelectedIndex = index
                            MyBase.OnSelectedIndexChanged(New EventArgs())
                            m_bTextChangedInternal = False
                        End If
                        Exit For
                    End If
                Next
                If index >= m_dvView.Count Then
                    m_bTextChangedInternal = True
                    m_iSelectedIndex = -1
                    SelectedIndex = -1
                    MyBase.OnSelectedIndexChanged(New EventArgs())
                    m_bTextChangedInternal = False
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message + vbCr & vbLf & "In ColumnComboBox.UpdateIndex().")
            End Try
        End Sub

        'Useful for setting the SelectedIndex to the index of a certain string.
        Public Function SetToIndexOf(sText As String) As Integer
            Try
                Dim index As Integer = 0
                'see if what is currently in the text box matches anything in the string list
                For index = 0 To m_slSuggestions.Count - 1
                    Dim sTemp As String = m_slSuggestions(index).ToUpper()
                    If sTemp = sText Then
                        Exit For
                    End If
                Next
                If index >= m_slSuggestions.Count Then
                    index = -1
                End If
                m_iSelectedIndex = index
                SelectedIndex = index
                MyBase.OnSelectedIndexChanged(New EventArgs())
                Return index
            Catch ex As Exception
                Throw New Exception(ex.Message + vbCr & vbLf & "In ColumnComboBox.SetToIndexOf(string).")
            End Try
        End Function



#Region ""

        'Overridden to do the set differently.
        Public Overrides Property Text() As String
            Get
                Return MyBase.Text
            End Get
            Set(value As String)
                'If MyBase.actualvalue Is Nothing Then
                If MyBase.Text <> value Then
                    MyBase.Text = value

                End If
            End Set
        End Property

        'Override the Selecte dindex to use the internal version.
        Public Overrides Property SelectedIndex() As Integer
            Get
                Return m_iSelectedIndex
            End Get
            Set(value As Integer)
                m_iSelectedIndex = value
                MyBase.SelectedIndex = value
                MyBase.OnSelectedIndexChanged(New EventArgs())
            End Set
        End Property

        'Property for getting and setting the DataTable that will be displayed in columns.
        Public Property Data() As DataTable
            Get
                Return m_dtData
            End Get
            Set(value As DataTable)
                If value Is Nothing Then
                    Throw New Exception("Data cannot be set to null." & vbCr & vbLf & " ColumnComboBox.Data (set)")
                End If
                m_dtData = value
                m_dvView = New DataView(m_dtData)
                m_bInitItems = True
                m_bInitSuggestionList = True
                Invalidate()
            End Set
        End Property

        'May be useful for getting the DataView used for displaying items
        Public Shadows ReadOnly Property Items() As DataView
            Get
                Return m_dvView
            End Get
        End Property

        'Access to the Columns so they can be hidden or shown or have widths set etc.
        Public ReadOnly Property Columns() As CCBColumnCollection
            Get
                If m_bInitItems Then
                    InitItems()
                End If
                If m_bInitDisplay Then
                    InitDisplay()
                End If
                Return m_Cols
            End Get
        End Property

        'Convenient for resorting the ComboBox based on a column.
        Public Sub SortBy(sCol As String, so As SortOrder)
            m_dvView.Sort = (sCol & Convert.ToString(" ")) + so.ToString()
            m_bInitItems = True
        End Sub

        'This is the column that will be displayed in the Text of the ComboBox and will also be used
        'for "suggesting" functionality.
        Public Property ViewColumn() As Integer
            Get
                Return m_iViewColumn
            End Get
            Set(value As Integer)
                If value < 0 Then
                    Throw New Exception("ViewColumn must be greater than zero" & vbCr & vbLf & "(set)ColumnComboBox.ViewColumn")
                End If
                m_iViewColumn = value
                m_bInitItems = True
                m_bInitDisplay = True
                m_bInitSuggestionList = True
            End Set
        End Property
        'Does nothing... yet
        Public Shadows ReadOnly Property Sorted() As Boolean
            Get
                Return False
            End Get
        End Property

        'Indexer for retriving values based on the column string.
        'Will return the value of the given column at SelectedIndex row.
        'You may want to add an int indexer as well.
        Default Public Property Item(sCol As String) As Object
            Get
                Try
                    If m_iSelectedIndex < 0 Then
                        Return Nothing
                    End If
                    Dim o As Object = Data.Rows(m_iSelectedIndex)(sCol)
                    Return o
                Catch ex As Exception
                    Throw New Exception(ex.Message + vbCr & vbLf & "In ColumnComboBox[string](get).")
                End Try
            End Get
            Set(value As Object)
                Try
                    Data.Rows(SelectedIndex)(sCol) = value
                Catch ex As Exception
                    Throw New Exception(ex.Message + vbCr & vbLf & "In ColumnComboBox[string](set).")
                End Try
            End Set
        End Property
#End Region

        'Event for changing which columns are displayed.
        Private Sub ColumnComboBox_OnColumnDisplayChanged(sender As Object, e As CCBColumnEventArgs)
            'Set the flag to re-init the display before next dropdown event
            m_bInitDisplay = True
        End Sub



    End Class
#Region "Supporting classes and enum"
    Public Enum SortOrder
        DESC
        ASC
    End Enum

    Public Class CCBColumnEventArgs
        Inherits EventArgs
        Public Column As CCBColumn
        Public Sub New(col As CCBColumn)
            Column = col
        End Sub
    End Class
    Public Delegate Sub ChangeColumnDisplayHandler(sender As Object, e As CCBColumnEventArgs)
    Public Class CCBColumn
        Private m_sName As String
        'public object Value;
        Public Width As Integer = -1
        Private m_Display As Boolean = True
        Public CalculatedWidth As Integer = 0

        Public Event OnColumnDisplayChanged As ChangeColumnDisplayHandler

        Public Sub New(sName As String)
            m_sName = sName
        End Sub
        Public Sub New(sName As String, iWidth As Integer)
            m_sName = sName
            Width = iWidth
        End Sub
        Public Sub New(sName As String, bDisplay As Boolean)
            m_sName = sName
            Display = bDisplay
        End Sub

#Region ""
        Public Property Name() As String
            Get
                Return m_sName
            End Get
            Set(value As String)
                If m_sName <> value Then
                    m_sName = value
                    RaiseEvent OnColumnDisplayChanged(Me, New CCBColumnEventArgs(Me))
                End If
            End Set
        End Property

        Public Property Display() As Boolean
            Get
                Return m_Display
            End Get
            Set(value As Boolean)
                If m_Display <> value Then
                    m_Display = value
                    RaiseEvent OnColumnDisplayChanged(Me, New CCBColumnEventArgs(Me))
                End If
            End Set
        End Property
#End Region
    End Class
    Public Class CCBColumnCollectionEventArgs
        Inherits EventArgs
        Public Count As Integer
        Public [DO] As CCBColumn
        Public Sub New(count__1 As Integer, dO__2 As CCBColumn)
            Count = count__1
            [DO] = dO__2
        End Sub
    End Class
    Public Delegate Sub AddCCBColumnHandler(sender As Object, e As CCBColumnCollectionEventArgs)
    Public Delegate Sub RemoveCCBColumnHandler(sender As Object, e As CCBColumnCollectionEventArgs)
    'CCBColumn collection is similar to an ArrayList but deals only with CCBColumns.
    'Sure would be nice to have class templates for classes like this one.
    Public Class CCBColumnCollection
        Implements IEnumerator

        Implements IEnumerable

        Private m_DOA As CCBColumn() = New CCBColumn(15) {}
        Private m_iSize As Integer = 16
        Private m_iCount As Integer = 0
        Private m_iEnumeratorPos As Integer
        Private m_bFireEvents As Boolean = True
        Public Event AddColumnEvent As AddCCBColumnHandler
        Public Event RemoveColumnEvent As RemoveCCBColumnHandler

        Public Sub New()
        End Sub
        Public Sub ItemAdded(sender As Object, e As CCBColumnCollectionEventArgs)
        End Sub
        Private Sub CheckGrow()
            If m_iCount >= m_iSize Then
                m_iSize *= 2
                Dim doTemp As CCBColumn() = New CCBColumn(m_iSize - 1) {}
                m_DOA.CopyTo(doTemp, 0)
                m_DOA = doTemp
            End If
        End Sub
        Protected Sub OnAddColumnEvent(e As CCBColumnCollectionEventArgs)
            RaiseEvent AddColumnEvent(Me, e)
        End Sub
        Public Sub Add([DO] As CCBColumn)
            If Contains([DO]) Then
                Throw New Exception((Convert.ToString("Column collection already contains a column named """) & [DO].Name) + """")
            End If
            CheckGrow()
            m_DOA(m_iCount) = [DO]
            m_iCount += 1

            Dim args As New CCBColumnCollectionEventArgs(m_iCount, [DO])
            RaiseEvent AddColumnEvent(Me, args)

            'If AddColumnEvent IsNot Nothing AndAlso m_bFireEvents Then
            '    Dim args As New CCBColumnCollectionEventArgs(m_iCount, [DO])
            '    RaiseEvent AddColumnEvent(Me, args)
            'End If
        End Sub
        Public Function Contains([DO] As CCBColumn) As Boolean
            For index As Integer = 0 To Count - 1
                If m_DOA(index).Name = [DO].Name Then
                    Return True
                End If
            Next
            Return False
        End Function
        Public Function AddNoDuplicate([DO] As CCBColumn) As Boolean
            Dim bRHS As Boolean = True
            If Contains([DO]) Then
                Remove([DO])
                bRHS = False
            End If
            Add([DO])
            Return bRHS
        End Function
        Public Sub Insert([DO] As CCBColumn, iPos As Integer)
            CheckGrow()
            If iPos < 0 Then
                Insert([DO], 0)
            End If
            If iPos >= m_iCount AndAlso iPos <> 0 Then
                Insert([DO], m_iCount - 1)
            End If
            Dim doTemp As CCBColumn() = New CCBColumn(m_iSize - 1) {}
            Dim index As Integer = 0
            While index < iPos
                doTemp(index) = m_DOA(index)
                index += 1
            End While
            doTemp(index) = [DO]
            While index < m_iCount
                doTemp(index + 1) = m_DOA(index)
                index += 1
            End While
            m_DOA = doTemp
            m_iCount += 1
        End Sub
        Public Sub Remove([DO] As CCBColumn)

            Dim index As Integer = 0
            While index < m_iCount
                If m_DOA(index).Name = [DO].Name Then
                    Exit While
                End If
                index += 1
            End While
            If index = m_iCount Then
                Return
            End If
            While index < m_iCount - 1
                m_DOA(index) = m_DOA(index + 1)
                index += 1
            End While
            m_iCount -= 1
            Remove([DO])
            Dim args As New CCBColumnCollectionEventArgs(m_iCount, [DO])
            RaiseEvent RemoveColumnEvent(Me, args)


            'If RemoveColumnEvent IsNot Nothing AndAlso m_bFireEvents Then
            '    Dim args As New CCBColumnCollectionEventArgs(m_iCount, [DO])
            '    RaiseEvent RemoveColumnEvent(Me, args)
            'End If
        End Sub
        Public Sub RemoveAt(index As Integer)
            If index < 0 OrElse index >= m_iCount Then
                Return
            End If
            While index < m_iCount - 1
                m_DOA(index) = m_DOA(index + 1)
                index += 1
            End While
            m_iCount -= 1
        End Sub
        Public Sub MoveToFront([DO] As CCBColumn)
            m_bFireEvents = False
            Remove([DO])
            Insert([DO], 0)
            m_bFireEvents = True
        End Sub
        Public Sub Clear()
            m_iSize = 16
            m_iCount = 0
            m_DOA = New CCBColumn(m_iSize - 1) {}
        End Sub
#Region ""
        Public ReadOnly Property Count() As Integer
            Get
                Return m_iCount
            End Get
        End Property

        Default Public ReadOnly Property Item(index As Integer) As CCBColumn
            Get
                Return m_DOA(index)
            End Get
        End Property

        Default Public ReadOnly Property Item(sName As String) As CCBColumn
            Get
                For index As Integer = 0 To m_iCount - 1
                    If m_DOA(index).Name = sName Then
                        Return m_DOA(index)
                    End If
                Next
                Throw New Exception((Convert.ToString("Column """) & sName) + """ is not a valid column.")
            End Get
        End Property
#End Region
#Region "IEnumerator Members"

        Public Function GetEnumerator() As IEnumerator
            m_iEnumeratorPos = -1
            Return DirectCast(Me, IEnumerator)
        End Function
        Public Sub Reset()
            m_iEnumeratorPos = -1
        End Sub

        Public ReadOnly Property Current() As Object
            Get
                Return m_DOA(m_iEnumeratorPos)
            End Get
        End Property

        Public Function MoveNext() As Boolean
            If m_iEnumeratorPos >= m_iCount - 1 Then
                Return False
            Else
                m_iEnumeratorPos += 1
                Return True
            End If
        End Function

#End Region

        Public ReadOnly Property Current1 As Object Implements System.Collections.IEnumerator.Current
            Get
                Return Nothing
            End Get
        End Property

        Public Function MoveNext1() As Boolean Implements System.Collections.IEnumerator.MoveNext

        End Function

        Public Sub Reset1() Implements System.Collections.IEnumerator.Reset

        End Sub

        Public Function GetEnumerator1() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
            Return Nothing
        End Function
    End Class

#End Region

End Namespace
