Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms

Friend Enum RunnerState
    Stopped
    Starting
    Running
    Paused
End Enum

Public Class MainForm
    Inherits Form
    Private Const MinIntervalMs As Integer = 50
    Private Const DefaultKeyHoldMs As Integer = 50
    Private Const DefaultStartDelaySeconds As Integer = 3
    Private Const HotKeyCaptureId As Integer = 1001
    Private Const HotKeyToggleId As Integer = 1002
    Private Const HotKeyStopId As Integer = 1003

    Private ReadOnly _points As List(Of ClickPointConfig)
    Private ReadOnly _pauseSignal As ManualResetEventSlim

    Private _currentIndex As Integer
    Private _runnerState As RunnerState
    Private _capturePending As Boolean
    Private _cts As CancellationTokenSource
    Private _clickTask As Task
    Private _startDelayCts As CancellationTokenSource

    Private _noteLabel As Label
    Private _pointsGrid As DataGridView
    Private _cmbActionType As ComboBox
    Private _txtX As TextBox
    Private _txtY As TextBox
    Private _cmbKey As ComboBox
    Private _txtHoldMs As TextBox
    Private _txtInterval As TextBox
    Private _txtRemark As TextBox
    Private _chkEnabled As CheckBox
    Private _btnAdd As Button
    Private _btnUpdate As Button
    Private _btnDelete As Button
    Private _btnMoveUp As Button
    Private _btnMoveDown As Button
    Private _btnCapture As Button
    Private _btnClear As Button
    Private _btnStart As Button
    Private _btnPause As Button
    Private _btnStop As Button
    Private _editorGroup As GroupBox
    Private _numStartDelay As NumericUpDown
    Private _chkMinimizeOnStart As CheckBox
    Private _runStateValue As Label
    Private _currentPointValue As Label
    Private _countdownValue As Label
    Private _lastClickValue As Label
    Private _captureValue As Label
    Private _configPathValue As Label

    Public Sub New()
        _points = New List(Of ClickPointConfig)()
        _pauseSignal = New ManualResetEventSlim(True)
        _currentIndex = -1
        _runnerState = RunnerState.Stopped
        _capturePending = False

        Me.Text = "前台连点器"
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.MinimumSize = New Size(960, 680)
        Me.Size = New Size(1040, 740)
        Me.Font = New Font("Microsoft YaHei UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)

        Dim rootLayout As New TableLayoutPanel()
        rootLayout.ColumnCount = 1
        rootLayout.RowCount = 5
        rootLayout.Dock = DockStyle.Fill
        rootLayout.Padding = New Padding(12)
        rootLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        Me.Controls.Add(rootLayout)

        _noteLabel = New Label()
        _noteLabel.AutoSize = True
        _noteLabel.ForeColor = Color.FromArgb(64, 64, 64)
        _noteLabel.Margin = New Padding(0, 0, 0, 10)
        _noteLabel.Text = "提示：执行时请保持目标窗口在前台。全屏应用抓不到 F9 时，请使用延迟启动并切回目标应用。"
        rootLayout.Controls.Add(_noteLabel, 0, 0)

        _pointsGrid = CreatePointsGrid()
        rootLayout.Controls.Add(_pointsGrid, 0, 1)

        _editorGroup = CreateEditorGroup()
        rootLayout.Controls.Add(_editorGroup, 0, 2)

        Dim controlPanel As FlowLayoutPanel = CreateControlPanel()
        rootLayout.Controls.Add(controlPanel, 0, 3)

        Dim statusGroup As GroupBox = CreateStatusGroup()
        rootLayout.Controls.Add(statusGroup, 0, 4)

        AddHandler Me.Load, AddressOf MainForm_Load
        AddHandler Me.FormClosing, AddressOf MainForm_FormClosing

        ClearEditor()
        LoadSavedConfig()
        RefreshGrid()
        UpdateUiState()
    End Sub

    Private Function CreatePointsGrid() As DataGridView
        Dim grid As New DataGridView()
        grid.Dock = DockStyle.Fill
        grid.AllowUserToAddRows = False
        grid.AllowUserToDeleteRows = False
        grid.AllowUserToResizeRows = False
        grid.MultiSelect = False
        grid.ReadOnly = True
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        grid.RowHeadersVisible = False
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        grid.BackgroundColor = Color.White
        grid.BorderStyle = BorderStyle.FixedSingle
        grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        grid.ColumnHeadersHeight = 34

        grid.Columns.Add(New DataGridViewTextBoxColumn() With {
            .Name = "Order",
            .HeaderText = "序号",
            .FillWeight = 10.0F
        })
        grid.Columns.Add(New DataGridViewTextBoxColumn() With {
            .Name = "ActionType",
            .HeaderText = "类型",
            .FillWeight = 14.0F
        })
        grid.Columns.Add(New DataGridViewTextBoxColumn() With {
            .Name = "X",
            .HeaderText = "X",
            .FillWeight = 10.0F
        })
        grid.Columns.Add(New DataGridViewTextBoxColumn() With {
            .Name = "Y",
            .HeaderText = "Y",
            .FillWeight = 10.0F
        })
        grid.Columns.Add(New DataGridViewTextBoxColumn() With {
            .Name = "Key",
            .HeaderText = "按键",
            .FillWeight = 14.0F
        })
        grid.Columns.Add(New DataGridViewTextBoxColumn() With {
            .Name = "Hold",
            .HeaderText = "按住(ms)",
            .FillWeight = 14.0F
        })
        grid.Columns.Add(New DataGridViewTextBoxColumn() With {
            .Name = "Interval",
            .HeaderText = "间隔(ms)",
            .FillWeight = 14.0F
        })
        grid.Columns.Add(New DataGridViewCheckBoxColumn() With {
            .Name = "Enabled",
            .HeaderText = "启用",
            .FillWeight = 10.0F,
            .ReadOnly = True
        })
        grid.Columns.Add(New DataGridViewTextBoxColumn() With {
            .Name = "Remark",
            .HeaderText = "备注",
            .FillWeight = 24.0F
        })

        AddHandler grid.SelectionChanged, AddressOf PointsGrid_SelectionChanged
        Return grid
    End Function

    Private Function CreateEditorGroup() As GroupBox
        Dim group As New GroupBox()
        group.Dock = DockStyle.Fill
        group.AutoSize = True
        group.Text = "动作编辑"
        group.Padding = New Padding(10)

        Dim layout As New TableLayoutPanel()
        layout.ColumnCount = 1
        layout.RowCount = 2
        layout.Dock = DockStyle.Fill
        layout.AutoSize = True
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        group.Controls.Add(layout)

        Dim fieldPanel As New FlowLayoutPanel()
        fieldPanel.Dock = DockStyle.Fill
        fieldPanel.AutoSize = True
        fieldPanel.WrapContents = True
        fieldPanel.Margin = New Padding(0)

        fieldPanel.Controls.Add(CreateFieldLabel("类型"))
        _cmbActionType = New ComboBox()
        _cmbActionType.DropDownStyle = ComboBoxStyle.DropDownList
        _cmbActionType.Width = 110
        _cmbActionType.Items.Add("鼠标左键")
        _cmbActionType.Items.Add("键盘按键")
        AddHandler _cmbActionType.SelectedIndexChanged, AddressOf ActionTypeChanged
        fieldPanel.Controls.Add(_cmbActionType)

        fieldPanel.Controls.Add(CreateFieldLabel("X"))
        _txtX = New TextBox()
        _txtX.Width = 70
        fieldPanel.Controls.Add(_txtX)

        fieldPanel.Controls.Add(CreateFieldLabel("Y"))
        _txtY = New TextBox()
        _txtY.Width = 70
        fieldPanel.Controls.Add(_txtY)

        fieldPanel.Controls.Add(CreateFieldLabel("按键"))
        _cmbKey = New ComboBox()
        _cmbKey.Width = 110
        _cmbKey.DropDownStyle = ComboBoxStyle.DropDown
        _cmbKey.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        _cmbKey.AutoCompleteSource = AutoCompleteSource.ListItems
        PopulateKeyOptions(_cmbKey)
        fieldPanel.Controls.Add(_cmbKey)

        fieldPanel.Controls.Add(CreateFieldLabel("按住(ms)"))
        _txtHoldMs = New TextBox()
        _txtHoldMs.Width = 80
        fieldPanel.Controls.Add(_txtHoldMs)

        fieldPanel.Controls.Add(CreateFieldLabel("间隔(ms)"))
        _txtInterval = New TextBox()
        _txtInterval.Width = 100
        fieldPanel.Controls.Add(_txtInterval)

        fieldPanel.Controls.Add(CreateFieldLabel("备注"))
        _txtRemark = New TextBox()
        _txtRemark.Width = 240
        fieldPanel.Controls.Add(_txtRemark)

        _chkEnabled = New CheckBox()
        _chkEnabled.AutoSize = True
        _chkEnabled.Checked = True
        _chkEnabled.Text = "启用该动作"
        _chkEnabled.Margin = New Padding(12, 8, 0, 0)
        fieldPanel.Controls.Add(_chkEnabled)

        layout.Controls.Add(fieldPanel, 0, 0)

        Dim buttonPanel As New FlowLayoutPanel()
        buttonPanel.Dock = DockStyle.Fill
        buttonPanel.AutoSize = True
        buttonPanel.Margin = New Padding(0, 10, 0, 0)

        _btnAdd = CreateActionButton("添加动作", AddressOf BtnAdd_Click)
        _btnUpdate = CreateActionButton("更新选中", AddressOf BtnUpdate_Click)
        _btnDelete = CreateActionButton("删除", AddressOf BtnDelete_Click)
        _btnMoveUp = CreateActionButton("上移", AddressOf BtnMoveUp_Click)
        _btnMoveDown = CreateActionButton("下移", AddressOf BtnMoveDown_Click)
        _btnCapture = CreateActionButton("开始取点", AddressOf BtnCapture_Click)
        _btnClear = CreateActionButton("清空输入", AddressOf BtnClear_Click)

        buttonPanel.Controls.Add(_btnAdd)
        buttonPanel.Controls.Add(_btnUpdate)
        buttonPanel.Controls.Add(_btnDelete)
        buttonPanel.Controls.Add(_btnMoveUp)
        buttonPanel.Controls.Add(_btnMoveDown)
        buttonPanel.Controls.Add(_btnCapture)
        buttonPanel.Controls.Add(_btnClear)

        layout.Controls.Add(buttonPanel, 0, 1)
        Return group
    End Function

    Private Function CreateControlPanel() As FlowLayoutPanel
        Dim panel As New FlowLayoutPanel()
        panel.Dock = DockStyle.Fill
        panel.AutoSize = True
        panel.Margin = New Padding(0, 12, 0, 12)

        panel.Controls.Add(CreateFieldLabel("启动延迟(s)"))
        _numStartDelay = New NumericUpDown()
        _numStartDelay.Width = 70
        _numStartDelay.Minimum = 0D
        _numStartDelay.Maximum = 15D
        _numStartDelay.Value = DefaultStartDelaySeconds
        panel.Controls.Add(_numStartDelay)

        _chkMinimizeOnStart = New CheckBox()
        _chkMinimizeOnStart.AutoSize = True
        _chkMinimizeOnStart.Checked = True
        _chkMinimizeOnStart.Margin = New Padding(8, 8, 16, 0)
        _chkMinimizeOnStart.Text = "启动后自动最小化"
        panel.Controls.Add(_chkMinimizeOnStart)

        _btnStart = CreateActionButton("开始", AddressOf BtnStart_Click)
        _btnPause = CreateActionButton("暂停", AddressOf BtnPause_Click)
        _btnStop = CreateActionButton("停止", AddressOf BtnStop_Click)

        _btnStart.Width = 100
        _btnPause.Width = 100
        _btnStop.Width = 100

        panel.Controls.Add(_btnStart)
        panel.Controls.Add(_btnPause)
        panel.Controls.Add(_btnStop)
        Return panel
    End Function

    Private Function CreateStatusGroup() As GroupBox
        Dim group As New GroupBox()
        group.Dock = DockStyle.Fill
        group.AutoSize = True
        group.Text = "运行状态"
        group.Padding = New Padding(10)

        Dim layout As New TableLayoutPanel()
        layout.ColumnCount = 2
        layout.RowCount = 6
        layout.Dock = DockStyle.Fill
        layout.AutoSize = True
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        group.Controls.Add(layout)

        _runStateValue = CreateStatusValueLabel("已停止")
        _currentPointValue = CreateStatusValueLabel("-")
        _countdownValue = CreateStatusValueLabel("0")
        _lastClickValue = CreateStatusValueLabel("-")
        _captureValue = CreateStatusValueLabel("按开始取点后使用 F8 记录当前鼠标位置。")
        _configPathValue = CreateStatusValueLabel(ConfigService.ConfigFilePath)

        layout.Controls.Add(CreateStatusTitleLabel("状态"), 0, 0)
        layout.Controls.Add(_runStateValue, 1, 0)
        layout.Controls.Add(CreateStatusTitleLabel("当前项"), 0, 1)
        layout.Controls.Add(_currentPointValue, 1, 1)
        layout.Controls.Add(CreateStatusTitleLabel("倒计时(ms)"), 0, 2)
        layout.Controls.Add(_countdownValue, 1, 2)
        layout.Controls.Add(CreateStatusTitleLabel("最近点击"), 0, 3)
        layout.Controls.Add(_lastClickValue, 1, 3)
        layout.Controls.Add(CreateStatusTitleLabel("取点"), 0, 4)
        layout.Controls.Add(_captureValue, 1, 4)
        layout.Controls.Add(CreateStatusTitleLabel("配置文件"), 0, 5)
        layout.Controls.Add(_configPathValue, 1, 5)

        Return group
    End Function

    Private Function CreateFieldLabel(text As String) As Label
        Dim label As New Label()
        label.AutoSize = True
        label.Text = text
        label.Margin = New Padding(0, 8, 6, 0)
        Return label
    End Function

    Private Function CreateActionButton(text As String, handler As EventHandler) As Button
        Dim button As New Button()
        button.AutoSize = True
        button.MinimumSize = New Size(92, 34)
        button.Text = text
        button.UseVisualStyleBackColor = True
        AddHandler button.Click, handler
        Return button
    End Function

    Private Function CreateStatusTitleLabel(text As String) As Label
        Dim label As New Label()
        label.AutoSize = True
        label.Margin = New Padding(0, 6, 12, 6)
        label.Font = New Font(Me.Font, FontStyle.Bold)
        label.Text = text
        Return label
    End Function

    Private Function CreateStatusValueLabel(text As String) As Label
        Dim label As New Label()
        label.AutoSize = True
        label.Margin = New Padding(0, 6, 0, 6)
        label.MaximumSize = New Size(860, 0)
        label.Text = text
        Return label
    End Function

    Private Sub PopulateKeyOptions(combo As ComboBox)
        Dim keys As String() = {
            "Space", "Enter", "Escape", "Tab",
            "Up", "Down", "Left", "Right",
            "W", "A", "S", "D",
            "J", "K", "L",
            "Q", "E", "R", "T",
            "F1", "F2", "F3", "F4", "F5", "F6",
            "F7", "F8", "F9", "F10", "F11", "F12",
            "D1", "D2", "D3", "D4", "D5",
            "NumPad1", "NumPad2", "NumPad3", "NumPad4", "NumPad5"
        }

        combo.Items.AddRange(keys)
    End Sub

    Private Function GetSelectedActionType() As ActionKind
        If _cmbActionType Is Nothing OrElse _cmbActionType.SelectedIndex <= 0 Then
            Return ActionKind.MouseClick
        End If

        Return ActionKind.KeyPress
    End Function

    Private Sub SetSelectedActionType(actionType As ActionKind)
        If _cmbActionType Is Nothing Then
            Return
        End If

        _cmbActionType.SelectedIndex = If(actionType = ActionKind.KeyPress, 1, 0)
    End Sub

    Private Sub ActionTypeChanged(sender As Object, e As EventArgs)
        UpdateEditorMode()
    End Sub

    Private Sub UpdateEditorMode()
        Dim actionType As ActionKind = GetSelectedActionType()
        Dim mouseMode As Boolean = (actionType = ActionKind.MouseClick)

        If _txtX IsNot Nothing Then
            _txtX.Enabled = mouseMode
        End If

        If _txtY IsNot Nothing Then
            _txtY.Enabled = mouseMode
        End If

        If _cmbKey IsNot Nothing Then
            _cmbKey.Enabled = Not mouseMode
        End If

        If _txtHoldMs IsNot Nothing Then
            _txtHoldMs.Enabled = Not mouseMode
        End If

        If _captureValue IsNot Nothing Then
            If mouseMode Then
                If _capturePending Then
                    _captureValue.Text = "取点模式已开启。请把鼠标移到目标位置，然后按 F8。"
                Else
                    _captureValue.Text = "鼠标动作可通过开始取点和 F8 录入屏幕坐标。"
                End If
            Else
                _capturePending = False
                _captureValue.Text = "键盘动作直接填写按键名称，例如 Space、Enter、W、A、S、D。"
            End If
        End If

        If _btnCapture IsNot Nothing Then
            _btnCapture.Enabled = (_runnerState = RunnerState.Stopped AndAlso mouseMode)
        End If
    End Sub

    Private Sub MainForm_Load(sender As Object, e As EventArgs)
        RegisterHotKeys()
    End Sub

    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs)
        UnregisterHotKeys()
        StopExecution(waitForCompletion:=True)
        SaveCurrentConfig(showError:=False)
    End Sub

    Private Sub RegisterHotKeys()
        Dim failures As New List(Of String)()

        RegisterSingleHotKey(HotKeyCaptureId, Keys.F8, "F8 取点", failures)
        RegisterSingleHotKey(HotKeyToggleId, Keys.F9, "F9 开始/暂停", failures)
        RegisterSingleHotKey(HotKeyStopId, Keys.F10, "F10 停止", failures)

        If failures.Count > 0 Then
            Dim details As String = String.Join("、", failures.ToArray())
            MessageBox.Show(Me, "以下热键注册失败：" & details & "。你仍可继续手动录入和点击。", "热键不可用", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            _captureValue.Text = "热键部分不可用，请改用手动录入。"
        End If
    End Sub

    Private Sub RegisterSingleHotKey(id As Integer, key As Keys, displayName As String, failures As List(Of String))
        If Not NativeMethods.RegisterHotKey(Me.Handle, id, 0UI, CUInt(key)) Then
            failures.Add(displayName)
        End If
    End Sub

    Private Sub UnregisterHotKeys()
        If Me.IsHandleCreated Then
            NativeMethods.UnregisterHotKey(Me.Handle, HotKeyCaptureId)
            NativeMethods.UnregisterHotKey(Me.Handle, HotKeyToggleId)
            NativeMethods.UnregisterHotKey(Me.Handle, HotKeyStopId)
        End If
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = NativeMethods.WM_HOTKEY Then
            Dim hotKeyId As Integer = m.WParam.ToInt32()

            Select Case hotKeyId
                Case HotKeyCaptureId
                    CaptureCurrentCursor()
                Case HotKeyToggleId
                    ToggleStartOrPause()
                Case HotKeyStopId
                    StopExecution(waitForCompletion:=False)
            End Select

            Return
        End If

        MyBase.WndProc(m)
    End Sub

    Private Sub BtnAdd_Click(sender As Object, e As EventArgs)
        Dim point As ClickPointConfig = Nothing
        If Not TryBuildPointFromEditor(point) Then
            Return
        End If

        _points.Add(point)
        RefreshGrid()
        SelectRow(_points.Count - 1)
        SaveCurrentConfig(showError:=False)
    End Sub

    Private Sub BtnUpdate_Click(sender As Object, e As EventArgs)
        Dim selectedIndex As Integer = GetSelectedIndex()
        If selectedIndex < 0 Then
            MessageBox.Show(Me, "请先选中需要更新的坐标。", "未选择项", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim point As ClickPointConfig = Nothing
        If Not TryBuildPointFromEditor(point) Then
            Return
        End If

        _points(selectedIndex) = point
        RefreshGrid()
        SelectRow(selectedIndex)
        SaveCurrentConfig(showError:=False)
    End Sub

    Private Sub BtnDelete_Click(sender As Object, e As EventArgs)
        Dim selectedIndex As Integer = GetSelectedIndex()
        If selectedIndex < 0 Then
            MessageBox.Show(Me, "请先选中需要删除的坐标。", "未选择项", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        _points.RemoveAt(selectedIndex)
        RefreshGrid()

        If _points.Count > 0 Then
            SelectRow(Math.Min(selectedIndex, _points.Count - 1))
        Else
            ClearEditor()
        End If

        SaveCurrentConfig(showError:=False)
    End Sub

    Private Sub BtnMoveUp_Click(sender As Object, e As EventArgs)
        Dim selectedIndex As Integer = GetSelectedIndex()
        If selectedIndex <= 0 Then
            Return
        End If

        Dim temp As ClickPointConfig = _points(selectedIndex - 1)
        _points(selectedIndex - 1) = _points(selectedIndex)
        _points(selectedIndex) = temp

        RefreshGrid()
        SelectRow(selectedIndex - 1)
        SaveCurrentConfig(showError:=False)
    End Sub

    Private Sub BtnMoveDown_Click(sender As Object, e As EventArgs)
        Dim selectedIndex As Integer = GetSelectedIndex()
        If selectedIndex < 0 OrElse selectedIndex >= _points.Count - 1 Then
            Return
        End If

        Dim temp As ClickPointConfig = _points(selectedIndex + 1)
        _points(selectedIndex + 1) = _points(selectedIndex)
        _points(selectedIndex) = temp

        RefreshGrid()
        SelectRow(selectedIndex + 1)
        SaveCurrentConfig(showError:=False)
    End Sub

    Private Sub BtnCapture_Click(sender As Object, e As EventArgs)
        SetSelectedActionType(ActionKind.MouseClick)
        _capturePending = True
        UpdateEditorMode()
    End Sub

    Private Sub BtnClear_Click(sender As Object, e As EventArgs)
        ClearEditor()
    End Sub

    Private Sub BtnStart_Click(sender As Object, e As EventArgs)
        StartExecution()
    End Sub

    Private Sub BtnPause_Click(sender As Object, e As EventArgs)
        If _runnerState = RunnerState.Running Then
            PauseExecution()
        ElseIf _runnerState = RunnerState.Paused Then
            ResumeExecution()
        End If
    End Sub

    Private Sub BtnStop_Click(sender As Object, e As EventArgs)
        StopExecution(waitForCompletion:=False)
    End Sub

    Private Sub PointsGrid_SelectionChanged(sender As Object, e As EventArgs)
        Dim selectedIndex As Integer = GetSelectedIndex()
        If selectedIndex < 0 OrElse selectedIndex >= _points.Count Then
            Return
        End If

        Dim point As ClickPointConfig = _points(selectedIndex)
        SetSelectedActionType(point.ActionType)
        _txtX.Text = point.X.ToString()
        _txtY.Text = point.Y.ToString()
        _cmbKey.Text = GetKeyName(point.KeyCode)
        _txtHoldMs.Text = point.HoldMs.ToString()
        _txtInterval.Text = point.IntervalMs.ToString()
        _txtRemark.Text = If(point.Remark, String.Empty)
        _chkEnabled.Checked = point.Enabled
        UpdateEditorMode()
        UpdateUiState()
    End Sub

    Private Sub ClearEditor()
        SetSelectedActionType(ActionKind.MouseClick)
        _txtX.Text = String.Empty
        _txtY.Text = String.Empty
        _cmbKey.Text = "Space"
        _txtHoldMs.Text = DefaultKeyHoldMs.ToString()
        _txtInterval.Text = "500"
        _txtRemark.Text = String.Empty
        _chkEnabled.Checked = True
        _capturePending = False
        UpdateEditorMode()
    End Sub

    Private Function TryBuildPointFromEditor(ByRef point As ClickPointConfig) As Boolean
        point = Nothing

        Dim intervalValue As Integer
        Dim xValue As Integer = 0
        Dim yValue As Integer = 0
        Dim holdValue As Integer = 0
        Dim parsedKey As Keys = Keys.None
        Dim actionType As ActionKind = GetSelectedActionType()

        If actionType = ActionKind.MouseClick Then
            If Not Integer.TryParse(_txtX.Text.Trim(), xValue) Then
                MessageBox.Show(Me, "X 坐标必须是整数。", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                _txtX.Focus()
                Return False
            End If

            If Not Integer.TryParse(_txtY.Text.Trim(), yValue) Then
                MessageBox.Show(Me, "Y 坐标必须是整数。", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                _txtY.Focus()
                Return False
            End If
        Else
            If Not TryParseKey(_cmbKey.Text.Trim(), parsedKey) Then
                MessageBox.Show(Me, "按键名称无效。可填写 Space、Enter、W、A、S、D、Left、Right 等。", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                _cmbKey.Focus()
                Return False
            End If

            If Not Integer.TryParse(_txtHoldMs.Text.Trim(), holdValue) Then
                MessageBox.Show(Me, "按住时长必须是整数。", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                _txtHoldMs.Focus()
                Return False
            End If

            If holdValue < 0 Then
                MessageBox.Show(Me, "按住时长不能小于 0 ms。", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                _txtHoldMs.Focus()
                Return False
            End If
        End If

        If Not Integer.TryParse(_txtInterval.Text.Trim(), intervalValue) Then
            MessageBox.Show(Me, "点击间隔必须是整数。", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            _txtInterval.Focus()
            Return False
        End If

        If intervalValue < MinIntervalMs Then
            MessageBox.Show(Me, "点击间隔不能小于 " & MinIntervalMs.ToString() & " ms。", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            _txtInterval.Focus()
            Return False
        End If

        point = New ClickPointConfig() With {
            .ActionType = actionType,
            .X = xValue,
            .Y = yValue,
            .IntervalMs = intervalValue,
            .Enabled = _chkEnabled.Checked,
            .Remark = _txtRemark.Text.Trim(),
            .KeyCode = CInt(parsedKey),
            .HoldMs = holdValue
        }
        Return True
    End Function

    Private Sub RefreshGrid()
        Dim previousSelection As Integer = GetSelectedIndex()

        _pointsGrid.Rows.Clear()
        For i As Integer = 0 To _points.Count - 1
            Dim point As ClickPointConfig = _points(i)
            _pointsGrid.Rows.Add(
                (i + 1).ToString(),
                GetActionTypeText(point.ActionType),
                If(point.ActionType = ActionKind.MouseClick, point.X.ToString(), "-"),
                If(point.ActionType = ActionKind.MouseClick, point.Y.ToString(), "-"),
                If(point.ActionType = ActionKind.KeyPress, GetKeyName(point.KeyCode), "-"),
                If(point.ActionType = ActionKind.KeyPress, point.HoldMs.ToString(), "-"),
                point.IntervalMs.ToString(),
                point.Enabled,
                If(point.Remark, String.Empty))
        Next

        If previousSelection >= 0 AndAlso previousSelection < _points.Count Then
            SelectRow(previousSelection)
        End If

        UpdateCurrentPointLabel()
        UpdateUiState()
    End Sub

    Private Function GetSelectedIndex() As Integer
        If _pointsGrid.SelectedRows.Count = 0 Then
            Return -1
        End If

        Return _pointsGrid.SelectedRows(0).Index
    End Function

    Private Sub SelectRow(index As Integer)
        If index < 0 OrElse index >= _pointsGrid.Rows.Count Then
            Return
        End If

        _pointsGrid.ClearSelection()
        _pointsGrid.Rows(index).Selected = True
        _pointsGrid.CurrentCell = _pointsGrid.Rows(index).Cells(0)
    End Sub

    Private Function ValidateBeforeStart() As Boolean
        If _points.Count = 0 Then
            MessageBox.Show(Me, "未配置任何动作，无法启动。", "无法启动", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return False
        End If

        Dim enabledCount As Integer = 0

        For i As Integer = 0 To _points.Count - 1
            Dim point As ClickPointConfig = _points(i)
            NormalizeAction(point)

            If point.IntervalMs < MinIntervalMs Then
                MessageBox.Show(Me, "第 " & (i + 1).ToString() & " 行的间隔无效。", "无法启动", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                SelectRow(i)
                Return False
            End If

            If point.ActionType = ActionKind.KeyPress AndAlso point.KeyCode <= 0 Then
                MessageBox.Show(Me, "第 " & (i + 1).ToString() & " 行的按键配置无效。", "无法启动", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                SelectRow(i)
                Return False
            End If

            If point.Enabled Then
                enabledCount += 1
            End If
        Next

        If enabledCount = 0 Then
            MessageBox.Show(Me, "所有动作均处于禁用状态，无法启动。", "无法启动", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return False
        End If

        Return True
    End Function

    Private Sub StartExecution()
        If _runnerState <> RunnerState.Stopped Then
            Return
        End If

        If Not ValidateBeforeStart() Then
            Return
        End If

        If _currentIndex < 0 OrElse _currentIndex >= _points.Count OrElse Not _points(_currentIndex).Enabled Then
            _currentIndex = FindFirstEnabledIndex()
        End If

        If _currentIndex < 0 Then
            MessageBox.Show(Me, "没有可执行的启用动作。", "无法启动", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        BeginStartDelay(CInt(_numStartDelay.Value))
    End Sub

    Private Sub BeginStartDelay(delaySeconds As Integer)
        CancelPendingStart()

        _capturePending = False
        _pauseSignal.Set()
        _runnerState = RunnerState.Starting
        _runStateValue.Text = If(delaySeconds > 0, "准备启动", "正在启动")
        UpdateCountdownLabel(delaySeconds * 1000)
        UpdateUiState()

        If _chkMinimizeOnStart.Checked Then
            Me.WindowState = FormWindowState.Minimized
        End If

        If delaySeconds <= 0 Then
            StartExecutionCore()
            Return
        End If

        _startDelayCts = New CancellationTokenSource()
        Dim token As CancellationToken = _startDelayCts.Token

        Task.Factory.StartNew(
            Sub() StartDelayLoop(delaySeconds, token),
            token,
            TaskCreationOptions.LongRunning,
            TaskScheduler.Default)
    End Sub

    Private Sub StartDelayLoop(delaySeconds As Integer, token As CancellationToken)
        Try
            Dim remainingMs As Integer = delaySeconds * 1000

            Do While remainingMs > 0
                token.ThrowIfCancellationRequested()
                UpdateCountdownLabel(remainingMs)

                Dim slice As Integer = Math.Min(100, remainingMs)
                If token.WaitHandle.WaitOne(slice) Then
                    token.ThrowIfCancellationRequested()
                End If

                remainingMs -= slice
            Loop

            SafeUi(AddressOf StartExecutionCore)
        Catch ex As OperationCanceledException
        End Try
    End Sub

    Private Sub StartExecutionCore()
        CancelPendingStart()

        If _runnerState <> RunnerState.Starting AndAlso _runnerState <> RunnerState.Stopped Then
            Return
        End If

        _pauseSignal.Set()
        _capturePending = False
        _runnerState = RunnerState.Running
        _runStateValue.Text = "运行中"
        _countdownValue.Text = "0"
        _clickTask = Nothing

        If _cts IsNot Nothing Then
            _cts.Dispose()
        End If

        _cts = New CancellationTokenSource()
        Dim token As CancellationToken = _cts.Token

        _clickTask = Task.Factory.StartNew(
            Sub() ClickLoop(token),
            token,
            TaskCreationOptions.LongRunning,
            TaskScheduler.Default)

        UpdateUiState()
    End Sub

    Private Sub PauseExecution()
        If _runnerState <> RunnerState.Running Then
            Return
        End If

        _pauseSignal.Reset()
        _runnerState = RunnerState.Paused
        _runStateValue.Text = "已暂停"
        UpdateUiState()
    End Sub

    Private Sub ResumeExecution()
        If _runnerState <> RunnerState.Paused Then
            Return
        End If

        _pauseSignal.Set()
        _runnerState = RunnerState.Running
        _runStateValue.Text = "运行中"
        UpdateUiState()
    End Sub

    Private Sub ToggleStartOrPause()
        If _runnerState = RunnerState.Stopped Then
            StartExecution()
        ElseIf _runnerState = RunnerState.Running Then
            PauseExecution()
        ElseIf _runnerState = RunnerState.Paused Then
            ResumeExecution()
        End If
    End Sub

    Private Sub StopExecution(waitForCompletion As Boolean)
        _capturePending = False
        CancelPendingStart()

        If _cts Is Nothing Then
            _runnerState = RunnerState.Stopped
            _runStateValue.Text = "已停止"
            UpdateCountdownLabel(0)
            UpdateUiState()
            Return
        End If

        _pauseSignal.Set()
        _runStateValue.Text = "正在停止"

        If Not _cts.IsCancellationRequested Then
            _cts.Cancel()
        End If

        If waitForCompletion AndAlso _clickTask IsNot Nothing Then
            Try
                _clickTask.Wait(1500)
            Catch ex As AggregateException
            Catch ex As OperationCanceledException
            End Try
        End If
    End Sub

    Private Sub CancelPendingStart()
        If _startDelayCts IsNot Nothing Then
            If Not _startDelayCts.IsCancellationRequested Then
                _startDelayCts.Cancel()
            End If

            _startDelayCts.Dispose()
            _startDelayCts = Nothing
        End If
    End Sub

    Private Sub ClickLoop(token As CancellationToken)
        Dim errorMessage As String = String.Empty

        Try
            Do While Not token.IsCancellationRequested
                _pauseSignal.Wait(token)

                If _currentIndex < 0 OrElse _currentIndex >= _points.Count Then
                    Exit Do
                End If

                Dim activeIndex As Integer = _currentIndex
                Dim point As ClickPointConfig = _points(activeIndex).Clone()
                If Not point.Enabled Then
                    _currentIndex = GetNextEnabledIndex(activeIndex)
                    If _currentIndex < 0 Then
                        Exit Do
                    End If
                    Continue Do
                End If

                SafeUi(
                    Sub()
                        _runStateValue.Text = If(_runnerState = RunnerState.Paused, "已暂停", "运行中")
                        _currentPointValue.Text = String.Format(
                            "{0}/{1} -> {2}  间隔 {3} ms",
                            (activeIndex + 1).ToString(),
                            _points.Count.ToString(),
                            GetActionSummary(point),
                            point.IntervalMs.ToString())
                    End Sub)

                ExecuteAction(point)

                SafeUi(
                    Sub()
                        _lastClickValue.Text = String.Format(
                            "{0} 执行 {1}",
                            DateTime.Now.ToString("HH:mm:ss"),
                            GetActionSummary(point))
                    End Sub)

                Dim nextIndex As Integer = GetNextEnabledIndex(activeIndex)
                If nextIndex < 0 Then
                    Exit Do
                End If

                _currentIndex = nextIndex
                WaitInterval(point.IntervalMs, token)
            Loop
        Catch ex As OperationCanceledException
        Catch ex As Exception
            errorMessage = ex.Message
        Finally
            SafeUi(Sub() FinishExecution(errorMessage))
        End Try
    End Sub

    Private Sub ExecuteAction(point As ClickPointConfig)
        If point.ActionType = ActionKind.KeyPress Then
            PerformKeyPress(point.KeyCode, point.HoldMs)
        Else
            PerformLeftClick(point.X, point.Y)
        End If
    End Sub

    Private Sub PerformLeftClick(x As Integer, y As Integer)
        SendAbsoluteMouseMove(x, y)
        Thread.Sleep(15)
        SendMouseFlags(NativeMethods.MOUSEEVENTF_LEFTDOWN)
        Thread.Sleep(15)
        SendMouseFlags(NativeMethods.MOUSEEVENTF_LEFTUP)
    End Sub

    Private Sub PerformKeyPress(keyCode As Integer, holdMs As Integer)
        Dim key As Keys = CType(keyCode, Keys)
        SendKeyInput(key, keyUp:=False)
        If holdMs > 0 Then
            Thread.Sleep(holdMs)
        End If
        SendKeyInput(key, keyUp:=True)
    End Sub

    Private Sub SendAbsoluteMouseMove(x As Integer, y As Integer)
        Dim virtualLeft As Integer = NativeMethods.GetSystemMetrics(NativeMethods.SM_XVIRTUALSCREEN)
        Dim virtualTop As Integer = NativeMethods.GetSystemMetrics(NativeMethods.SM_YVIRTUALSCREEN)
        Dim virtualWidth As Integer = NativeMethods.GetSystemMetrics(NativeMethods.SM_CXVIRTUALSCREEN)
        Dim virtualHeight As Integer = NativeMethods.GetSystemMetrics(NativeMethods.SM_CYVIRTUALSCREEN)

        If virtualWidth <= 0 Then
            virtualWidth = Screen.PrimaryScreen.Bounds.Width
            virtualLeft = 0
        End If

        If virtualHeight <= 0 Then
            virtualHeight = Screen.PrimaryScreen.Bounds.Height
            virtualTop = 0
        End If

        Dim normalizedX As Integer = NormalizeAbsoluteCoordinate(x, virtualLeft, virtualWidth)
        Dim normalizedY As Integer = NormalizeAbsoluteCoordinate(y, virtualTop, virtualHeight)

        Dim inputItem As New NativeMethods.INPUT()
        inputItem.type = NativeMethods.INPUT_MOUSE
        inputItem.unionData.mi = New NativeMethods.MOUSEINPUT() With {
            .dx = normalizedX,
            .dy = normalizedY,
            .mouseData = 0UI,
            .dwFlags = NativeMethods.MOUSEEVENTF_MOVE Or NativeMethods.MOUSEEVENTF_ABSOLUTE Or NativeMethods.MOUSEEVENTF_VIRTUALDESK,
            .time = 0UI,
            .dwExtraInfo = IntPtr.Zero
        }

        SendInputOrThrow(New NativeMethods.INPUT() {inputItem})
    End Sub

    Private Function NormalizeAbsoluteCoordinate(value As Integer, origin As Integer, length As Integer) As Integer
        If length <= 1 Then
            Return 0
        End If

        Dim adjusted As Double = CDbl(value - origin)
        Dim scaled As Double = adjusted * 65535.0R / CDbl(length - 1)
        Return CInt(Math.Max(0.0R, Math.Min(65535.0R, Math.Round(scaled))))
    End Function

    Private Sub SendMouseFlags(flags As UInteger)
        Dim inputItem As New NativeMethods.INPUT()
        inputItem.type = NativeMethods.INPUT_MOUSE
        inputItem.unionData.mi = New NativeMethods.MOUSEINPUT() With {
            .dx = 0,
            .dy = 0,
            .mouseData = 0UI,
            .dwFlags = flags,
            .time = 0UI,
            .dwExtraInfo = IntPtr.Zero
        }

        SendInputOrThrow(New NativeMethods.INPUT() {inputItem})
    End Sub

    Private Sub SendKeyInput(key As Keys, keyUp As Boolean)
        Dim keyCode As Integer = CInt(key) And &HFFFF
        Dim scanCode As UInteger = NativeMethods.MapVirtualKey(CUInt(keyCode), NativeMethods.MAPVK_VK_TO_VSC)
        Dim flags As UInteger = 0UI

        If scanCode <> 0UI Then
            flags = flags Or NativeMethods.KEYEVENTF_SCANCODE
        End If

        If IsExtendedKey(key) Then
            flags = flags Or NativeMethods.KEYEVENTF_EXTENDEDKEY
        End If

        If keyUp Then
            flags = flags Or NativeMethods.KEYEVENTF_KEYUP
        End If

        Dim virtualKeyValue As UShort = If(scanCode <> 0UI, CUShort(0), CUShort(keyCode And &HFFFF))

        Dim inputItem As New NativeMethods.INPUT()
        inputItem.type = NativeMethods.INPUT_KEYBOARD
        inputItem.unionData.ki = New NativeMethods.KEYBDINPUT() With {
            .wVk = virtualKeyValue,
            .wScan = CUShort(scanCode And &HFFFFUI),
            .dwFlags = flags,
            .time = 0UI,
            .dwExtraInfo = IntPtr.Zero
        }

        SendInputOrThrow(New NativeMethods.INPUT() {inputItem})
    End Sub

    Private Sub SendInputOrThrow(inputs As NativeMethods.INPUT())
        Dim expected As UInteger = CUInt(inputs.Length)
        Dim actual As UInteger = NativeMethods.SendInput(expected, inputs, Marshal.SizeOf(GetType(NativeMethods.INPUT)))
        If actual <> expected Then
            Throw New InvalidOperationException("SendInput 未能完整发送输入。")
        End If
    End Sub

    Private Sub WaitInterval(totalMs As Integer, token As CancellationToken)
        Dim remaining As Integer = totalMs
        UpdateCountdownLabel(remaining)

        Do While remaining > 0
            _pauseSignal.Wait(token)

            Dim slice As Integer = Math.Min(50, remaining)
            If token.WaitHandle.WaitOne(slice) Then
                token.ThrowIfCancellationRequested()
            End If

            remaining -= slice
            UpdateCountdownLabel(remaining)
        Loop
    End Sub

    Private Sub FinishExecution(errorMessage As String)
        If _cts IsNot Nothing Then
            _cts.Dispose()
            _cts = Nothing
        End If

        _clickTask = Nothing
        _pauseSignal.Set()
        _runnerState = RunnerState.Stopped
        _currentIndex = FindFirstEnabledIndex()
        _countdownValue.Text = "0"
        UpdateCurrentPointLabel()
        UpdateUiState()

        If String.IsNullOrEmpty(errorMessage) Then
            _runStateValue.Text = "已停止"
        Else
            _runStateValue.Text = "异常停止"
            MessageBox.Show(Me, "执行过程中出现错误：" & errorMessage, "执行失败", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub CaptureCurrentCursor()
        If Not _capturePending Then
            Return
        End If

        Dim position As Point = Cursor.Position
        SafeUi(
            Sub()
                _txtX.Text = position.X.ToString()
                _txtY.Text = position.Y.ToString()
                _capturePending = False
                _captureValue.Text = String.Format("已记录坐标 ({0}, {1})。", position.X.ToString(), position.Y.ToString())
            End Sub)
    End Sub

    Private Function FindFirstEnabledIndex() As Integer
        For i As Integer = 0 To _points.Count - 1
            If _points(i).Enabled Then
                Return i
            End If
        Next

        Return -1
    End Function

    Private Function GetNextEnabledIndex(currentIndex As Integer) As Integer
        If _points.Count = 0 Then
            Return -1
        End If

        If currentIndex < 0 Then
            Return FindFirstEnabledIndex()
        End If

        For offset As Integer = 1 To _points.Count
            Dim candidate As Integer = (currentIndex + offset) Mod _points.Count
            If _points(candidate).Enabled Then
                Return candidate
            End If
        Next

        Return -1
    End Function

    Private Sub NormalizeAction(point As ClickPointConfig)
        If point Is Nothing Then
            Return
        End If

        If point.Remark Is Nothing Then
            point.Remark = String.Empty
        End If

        If point.ActionType = ActionKind.KeyPress Then
            If point.HoldMs < 0 Then
                point.HoldMs = DefaultKeyHoldMs
            End If
        Else
            point.KeyCode = 0
            point.HoldMs = 0
        End If
    End Sub

    Private Function TryParseKey(keyText As String, ByRef parsedKey As Keys) As Boolean
        parsedKey = Keys.None

        If String.IsNullOrWhiteSpace(keyText) Then
            Return False
        End If

        Dim normalized As String = keyText.Trim()

        If normalized.Length = 1 AndAlso Char.IsDigit(normalized(0)) Then
            normalized = "D" & normalized
        End If

        Try
            Dim candidate As Keys = CType([Enum].Parse(GetType(Keys), normalized, True), Keys)
            If candidate = Keys.None Then
                Return False
            End If

            parsedKey = candidate
            Return True
        Catch ex As ArgumentException
            Return False
        End Try
    End Function

    Private Function GetKeyName(keyCode As Integer) As String
        If keyCode <= 0 Then
            Return String.Empty
        End If

        Return CType(keyCode, Keys).ToString()
    End Function

    Private Function GetActionTypeText(actionType As ActionKind) As String
        If actionType = ActionKind.KeyPress Then
            Return "按键"
        End If

        Return "鼠标"
    End Function

    Private Function GetActionSummary(point As ClickPointConfig) As String
        If point.ActionType = ActionKind.KeyPress Then
            Return String.Format("按键 {0} 按住 {1} ms", GetKeyName(point.KeyCode), point.HoldMs.ToString())
        End If

        Return String.Format("鼠标点击 ({0}, {1})", point.X.ToString(), point.Y.ToString())
    End Function

    Private Function IsExtendedKey(key As Keys) As Boolean
        Select Case key
            Case Keys.Up, Keys.Down, Keys.Left, Keys.Right, Keys.Home, Keys.End, Keys.PageUp, Keys.PageDown, Keys.Insert, Keys.Delete, Keys.NumLock, Keys.PrintScreen
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Sub UpdateCurrentPointLabel()
        If _currentIndex >= 0 AndAlso _currentIndex < _points.Count Then
            Dim point As ClickPointConfig = _points(_currentIndex)
            _currentPointValue.Text = String.Format("{0}/{1} -> {2}", (_currentIndex + 1).ToString(), _points.Count.ToString(), GetActionSummary(point))
        Else
            _currentPointValue.Text = "-"
        End If
    End Sub

    Private Sub UpdateCountdownLabel(value As Integer)
        SafeUi(Sub() _countdownValue.Text = value.ToString())
    End Sub

    Private Sub UpdateUiState()
        Dim isStopped As Boolean = (_runnerState = RunnerState.Stopped)
        Dim isStarting As Boolean = (_runnerState = RunnerState.Starting)
        Dim selectedIndex As Integer = GetSelectedIndex()

        _editorGroup.Enabled = isStopped
        _pointsGrid.Enabled = isStopped

        _btnAdd.Enabled = isStopped
        _btnUpdate.Enabled = isStopped AndAlso selectedIndex >= 0
        _btnDelete.Enabled = isStopped AndAlso selectedIndex >= 0
        _btnMoveUp.Enabled = isStopped AndAlso selectedIndex > 0
        _btnMoveDown.Enabled = isStopped AndAlso selectedIndex >= 0 AndAlso selectedIndex < _points.Count - 1
        _btnClear.Enabled = isStopped
        _numStartDelay.Enabled = isStopped
        _chkMinimizeOnStart.Enabled = isStopped

        _btnStart.Enabled = isStopped
        _btnPause.Enabled = (_runnerState = RunnerState.Running OrElse _runnerState = RunnerState.Paused)
        _btnStop.Enabled = (Not isStopped OrElse isStarting)
        _btnPause.Text = If(_runnerState = RunnerState.Paused, "继续", "暂停")
        UpdateEditorMode()
    End Sub

    Private Sub SaveCurrentConfig(showError As Boolean)
        Dim config As New AppConfig()
        Dim bounds As Rectangle = If(Me.WindowState = FormWindowState.Normal, Me.Bounds, Me.RestoreBounds)

        config.WindowX = bounds.X
        config.WindowY = bounds.Y
        config.WindowWidth = bounds.Width
        config.WindowHeight = bounds.Height
        config.StartDelaySeconds = CInt(_numStartDelay.Value)
        config.MinimizeOnStart = _chkMinimizeOnStart.Checked

        For Each point As ClickPointConfig In _points
            config.Points.Add(point.Clone())
        Next

        Dim errorMessage As String = String.Empty
        If Not ConfigService.SaveConfig(config, errorMessage) AndAlso showError Then
            MessageBox.Show(Me, "配置保存失败：" & errorMessage, "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub LoadSavedConfig()
        Dim hadError As Boolean = False
        Dim errorMessage As String = String.Empty
        Dim config As AppConfig = ConfigService.LoadConfig(hadError, errorMessage)

        _points.Clear()
        For Each point As ClickPointConfig In config.Points
            NormalizeAction(point)
            _points.Add(point.Clone())
        Next

        If config.WindowWidth > 0 AndAlso config.WindowHeight > 0 Then
            Me.StartPosition = FormStartPosition.Manual
            Me.Bounds = New Rectangle(config.WindowX, config.WindowY, config.WindowWidth, config.WindowHeight)
        End If

        Dim delayValue As Integer = config.StartDelaySeconds
        If delayValue < 0 Then
            delayValue = DefaultStartDelaySeconds
        ElseIf delayValue > CInt(_numStartDelay.Maximum) Then
            delayValue = CInt(_numStartDelay.Maximum)
        End If

        _numStartDelay.Value = delayValue
        _chkMinimizeOnStart.Checked = config.MinimizeOnStart

        _currentIndex = FindFirstEnabledIndex()
        _configPathValue.Text = ConfigService.ConfigFilePath

        If hadError Then
            MessageBox.Show(Me, "配置文件读取失败，程序已回退为空列表。" & Environment.NewLine & errorMessage, "配置读取失败", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub SafeUi(action As Action)
        If Me.IsDisposed Then
            Return
        End If

        If Me.InvokeRequired Then
            Try
                Me.BeginInvoke(action)
            Catch ex As InvalidOperationException
            End Try
        Else
            action()
        End If
    End Sub
End Class
