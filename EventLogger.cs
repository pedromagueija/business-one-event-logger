// --------------------------------------------------------------------------------------------------------------------
// <copyright file="EventLogger.cs" company="">
//   
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace EventLogger
{
    using System;
    using System.ComponentModel;
    using System.Data;
    using System.Drawing;
    using System.IO;
    using System.Windows.Forms;

    using SAPbouiCOM;

    using Application = SAPbouiCOM.Application;
    using DataColumn = System.Data.DataColumn;
    using DataTable = System.Data.DataTable;
    using Form = System.Windows.Forms.Form;

    /// <summary>
    /// The event logger.
    /// </summary>
    public class EventLogger : Form
    {
        /// <summary>
        /// The the appl.
        /// </summary>
        private Application theAppl;

        /// <summary>
        /// The the gui.
        /// </summary>
        private SboGuiApi theGui;

        /// <summary>
        /// The add events log method.
        /// </summary>
        /// <param name="eventInfo">
        /// The event info.
        /// </param>
        /// <param name="typeEvent">
        /// The type event.
        /// </param>
        /// <param name="before">
        /// The before.
        /// </param>
        /// <param name="formType">
        /// The form type.
        /// </param>
        /// <param name="formCount">
        /// The form count.
        /// </param>
        /// <param name="itemUID">
        /// The item uid.
        /// </param>
        /// <param name="colUID">
        /// The col uid.
        /// </param>
        /// <param name="rowNb">
        /// The row nb.
        /// </param>
        /// <param name="formUID">
        /// The form uid.
        /// </param>
        /// <param name="formMode">
        /// The form mode.
        /// </param>
        /// <param name="charPressed">
        /// The char pressed.
        /// </param>
        /// <param name="innerEvent">
        /// The inner event.
        /// </param>
        /// <param name="itemChanged">
        /// The item changed.
        /// </param>
        /// <param name="modifiers">
        /// The modifiers.
        /// </param>
        /// <param name="popUpIndicator">
        /// The pop up indicator.
        /// </param>
        /// <param name="success">
        /// The success.
        /// </param>
        private delegate void AddEventsLogMethod(string eventInfo, string typeEvent, string before, string formType, string formCount, string itemUID, string colUID, string rowNb, string formUID, string formMode, string charPressed, string innerEvent, string itemChanged, string modifiers, string popUpIndicator, string success);

        /// <summary>
        /// The evt log.
        /// </summary>
        private readonly AddEventsLogMethod evtLog;

        /// <summary>
        /// The filter file.
        /// </summary>
        public static string FilterFile = "evtLog_Filter.txt";

        /// <summary>
        /// The separator.
        /// </summary>
        public static string Separator = ";";

        /// <summary>
        /// The data grid.
        /// </summary>
        private DataGrid dataGrid;

        /// <summary>
        /// The components.
        /// </summary>
        private IContainer components;

        /// <summary>
        /// The data table.
        /// </summary>
        private DataTable dataTable;

        /// <summary>
        /// The nb col.
        /// </summary>
        private DataColumn nbCol;

        /// <summary>
        /// The time col.
        /// </summary>
        private DataColumn timeCol;

        /// <summary>
        /// The event col.
        /// </summary>
        private DataColumn eventCol;

        /// <summary>
        /// The type col.
        /// </summary>
        private DataColumn typeCol;

        /// <summary>
        /// The before col.
        /// </summary>
        private DataColumn beforeCol;

        /// <summary>
        /// The form type col.
        /// </summary>
        private DataColumn formTypeCol;

        /// <summary>
        /// The form count col.
        /// </summary>
        private DataColumn formCountCol;

        /// <summary>
        /// The item uid col.
        /// </summary>
        private DataColumn itemUIDCol;

        /// <summary>
        /// The col uid col.
        /// </summary>
        private DataColumn colUIDCol;

        /// <summary>
        /// The row col.
        /// </summary>
        private DataColumn rowCol;

        /// <summary>
        /// The form uid col.
        /// </summary>
        private DataColumn formUIDCol;

        /// <summary>
        /// The form mode col.
        /// </summary>
        private DataColumn formModeCol;

        /// <summary>
        /// The char pressed col.
        /// </summary>
        private DataColumn charPressedCol;

        /// <summary>
        /// The inner event col.
        /// </summary>
        private DataColumn innerEventCol;

        /// <summary>
        /// The item changed col.
        /// </summary>
        private DataColumn itemChangedCol;

        /// <summary>
        /// The modifiers col.
        /// </summary>
        private DataColumn modifiersCol;

        /// <summary>
        /// The pop up indicator col.
        /// </summary>
        private DataColumn popUpIndicatorCol;

        /// <summary>
        /// The success col.
        /// </summary>
        private DataColumn successCol;

        /// <summary>
        /// The data set.
        /// </summary>
        private DataSet dataSet;

        /// <summary>
        /// The tool bar 1.
        /// </summary>
        private ToolBar toolBar1;

        /// <summary>
        /// The connect_tool bar button.
        /// </summary>
        private ToolBarButton connect_toolBarButton;

        /// <summary>
        /// The disconnect_tool bar button.
        /// </summary>
        private ToolBarButton disconnect_toolBarButton;

        /// <summary>
        /// The tool bar button 1.
        /// </summary>
        private ToolBarButton toolBarButton1;

        /// <summary>
        /// The tool bar button 2.
        /// </summary>
        private ToolBarButton toolBarButton2;

        /// <summary>
        /// The tool bar button 3.
        /// </summary>
        private ToolBarButton toolBarButton3;

        /// <summary>
        /// The tool bar button 4.
        /// </summary>
        private ToolBarButton toolBarButton4;

        /// <summary>
        /// The tool bar button 5.
        /// </summary>
        private ToolBarButton toolBarButton5;

        /// <summary>
        /// The exit_tool bar button.
        /// </summary>
        private ToolBarButton exit_toolBarButton;

        /// <summary>
        /// The tool bar button 6.
        /// </summary>
        private ToolBarButton toolBarButton6;

        /// <summary>
        /// The notify icon 1.
        /// </summary>
        private NotifyIcon notifyIcon1;

        /// <summary>
        /// The clear log_tool bar button.
        /// </summary>
        private ToolBarButton clearLog_toolBarButton;

        /// <summary>
        /// The filter_tool bar button.
        /// </summary>
        private ToolBarButton filter_toolBarButton;

        /// <summary>
        /// The image list 1.
        /// </summary>
        private ImageList imageList1;

        /// <summary>
        /// Initializes a new instance of the <see cref="EventLogger"/> class.
        /// </summary>
        public EventLogger()
        {
            this.InitializeComponent();

            this.InitializeGrid();

            this.evtLog = this.addEventsLog;

            if (this.Connect())
            {
                this.connect_toolBarButton.Enabled = false;
                this.disconnect_toolBarButton.Enabled = true;
            }
            else
            {
                this.connect_toolBarButton.Enabled = true;
                this.disconnect_toolBarButton.Enabled = false;
            }

            this.notifyIcon1.Icon = new Icon("icons\\property.ico");
            this.notifyIcon1.Text = "EventLogger";
            this.notifyIcon1.Visible = true;
        }

        #region Windows Form Designer 

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new Container();
            ComponentResourceManager resources = new ComponentResourceManager(typeof(EventLogger));
            this.dataGrid = new DataGrid();
            this.toolBar1 = new ToolBar();
            this.connect_toolBarButton = new ToolBarButton();
            this.disconnect_toolBarButton = new ToolBarButton();
            this.toolBarButton1 = new ToolBarButton();
            this.toolBarButton2 = new ToolBarButton();
            this.filter_toolBarButton = new ToolBarButton();
            this.clearLog_toolBarButton = new ToolBarButton();
            this.toolBarButton3 = new ToolBarButton();
            this.toolBarButton4 = new ToolBarButton();
            this.toolBarButton5 = new ToolBarButton();
            this.toolBarButton6 = new ToolBarButton();
            this.exit_toolBarButton = new ToolBarButton();
            this.imageList1 = new ImageList(this.components);
            this.notifyIcon1 = new NotifyIcon(this.components);
            ((ISupportInitialize)this.dataGrid).BeginInit();
            this.SuspendLayout();

            // dataGrid
            this.dataGrid.Anchor = ((AnchorStyles.Top | AnchorStyles.Bottom) | AnchorStyles.Left) | AnchorStyles.Right;
            this.dataGrid.DataMember = string.Empty;
            this.dataGrid.HeaderForeColor = SystemColors.ControlText;
            this.dataGrid.Location = new Point(0, 32);
            this.dataGrid.Name = "dataGrid";
            this.dataGrid.Size = new Size(600, 552);
            this.dataGrid.TabIndex = 0;

            // toolBar1
            this.toolBar1.Buttons.AddRange(new[] { this.connect_toolBarButton, this.disconnect_toolBarButton, this.toolBarButton1, this.toolBarButton2, this.filter_toolBarButton, this.clearLog_toolBarButton, this.toolBarButton3, this.toolBarButton4, this.toolBarButton5, this.toolBarButton6, this.exit_toolBarButton });
            this.toolBar1.DropDownArrows = true;
            this.toolBar1.ImageList = this.imageList1;
            this.toolBar1.Location = new Point(0, 0);
            this.toolBar1.Name = "toolBar1";
            this.toolBar1.ShowToolTips = true;
            this.toolBar1.Size = new Size(602, 28);
            this.toolBar1.TabIndex = 5;
            this.toolBar1.ButtonClick += this.toolBar1_ButtonClick;

            // connect_toolBarButton
            this.connect_toolBarButton.ImageIndex = 0;
            this.connect_toolBarButton.Name = "connect_toolBarButton";
            this.connect_toolBarButton.ToolTipText = "Connect";

            // disconnect_toolBarButton
            this.disconnect_toolBarButton.ImageIndex = 2;
            this.disconnect_toolBarButton.Name = "disconnect_toolBarButton";
            this.disconnect_toolBarButton.ToolTipText = "Disconnect";

            // toolBarButton1
            this.toolBarButton1.Name = "toolBarButton1";
            this.toolBarButton1.Style = ToolBarButtonStyle.Separator;

            // toolBarButton2
            this.toolBarButton2.Name = "toolBarButton2";
            this.toolBarButton2.Style = ToolBarButtonStyle.Separator;

            // filter_toolBarButton
            this.filter_toolBarButton.ImageIndex = 6;
            this.filter_toolBarButton.Name = "filter_toolBarButton";
            this.filter_toolBarButton.ToolTipText = "Filter Events";

            // clearLog_toolBarButton
            this.clearLog_toolBarButton.ImageIndex = 1;
            this.clearLog_toolBarButton.Name = "clearLog_toolBarButton";
            this.clearLog_toolBarButton.ToolTipText = "ClearLog";

            // toolBarButton3
            this.toolBarButton3.Name = "toolBarButton3";
            this.toolBarButton3.Style = ToolBarButtonStyle.Separator;

            // toolBarButton4
            this.toolBarButton4.Name = "toolBarButton4";
            this.toolBarButton4.Style = ToolBarButtonStyle.Separator;

            // toolBarButton5
            this.toolBarButton5.Name = "toolBarButton5";
            this.toolBarButton5.Style = ToolBarButtonStyle.Separator;

            // toolBarButton6
            this.toolBarButton6.Name = "toolBarButton6";
            this.toolBarButton6.Style = ToolBarButtonStyle.Separator;

            // exit_toolBarButton
            this.exit_toolBarButton.ImageIndex = 3;
            this.exit_toolBarButton.Name = "exit_toolBarButton";
            this.exit_toolBarButton.ToolTipText = "Exit";

            // imageList1
            this.imageList1.ImageStream = (ImageListStreamer)(resources.GetObject("imageList1.ImageStream"));
            this.imageList1.TransparentColor = Color.Transparent;
            this.imageList1.Images.SetKeyName(0, string.Empty);
            this.imageList1.Images.SetKeyName(1, string.Empty);
            this.imageList1.Images.SetKeyName(2, string.Empty);
            this.imageList1.Images.SetKeyName(3, string.Empty);
            this.imageList1.Images.SetKeyName(4, string.Empty);
            this.imageList1.Images.SetKeyName(5, string.Empty);
            this.imageList1.Images.SetKeyName(6, string.Empty);

            // notifyIcon1
            this.notifyIcon1.Text = "notifyIcon1";
            this.notifyIcon1.Visible = true;
            this.notifyIcon1.DoubleClick += this.notifyIcon1_MouseDoubleClick;

            // EventLogger
            this.AutoScaleBaseSize = new Size(5, 13);
            this.ClientSize = new Size(602, 586);
            this.Controls.Add(this.toolBar1);
            this.Controls.Add(this.dataGrid);
            this.Icon = (Icon)(resources.GetObject("$this.Icon"));
            this.Name = "EventLogger";
            this.Text = "Event Logger";
            this.Closed += this.CloseHandler;
            this.Resize += this.ResizeHandler;
            ((ISupportInitialize)this.dataGrid).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        /// <summary>
        /// The initialize grid.
        /// </summary>
        protected void InitializeGrid()
        {
            // Check B1 Version is correct regarding B1TE version installed
            // B1TEUtils.B1Version.CheckB1Version();
            this.connect_toolBarButton.Enabled = false;
            this.clearLog_toolBarButton.Enabled = false;

            /////////////////////////////////////////////////////////////////
            this.dataTable = new DataTable("EventLoggerData");

            /////////////////////////////////////////////////////////////////
            this.nbCol = new DataColumn("#", typeof(string));
            this.dataTable.Columns.Add(this.nbCol);

            DataGridTextBoxColumn nbColumnStyle = new DataGridTextBoxColumn();
            nbColumnStyle.MappingName = "#";
            nbColumnStyle.HeaderText = "#";
            nbColumnStyle.Width = 25;

            /////////////////////////////////////////////////////////////////
            this.timeCol = new DataColumn("Time", typeof(string));
            this.dataTable.Columns.Add(this.timeCol);

            DataGridTextBoxColumn timeColumnStyle = new DataGridTextBoxColumn();
            timeColumnStyle.MappingName = "Time";
            timeColumnStyle.HeaderText = "Time";
            timeColumnStyle.Width = 80;

            /////////////////////////////////////////////////////////////////
            this.eventCol = new DataColumn("Event", typeof(string));
            this.dataTable.Columns.Add(this.eventCol);

            DataGridTextBoxColumn eventColumnStyle = new DataGridTextBoxColumn();
            eventColumnStyle.MappingName = "Event";
            eventColumnStyle.HeaderText = "Event";
            eventColumnStyle.Width = 70;

            /////////////////////////////////////////////////////////////////
            this.typeCol = new DataColumn("EvtType", typeof(string));
            this.dataTable.Columns.Add(this.typeCol);

            DataGridTextBoxColumn typeColumnStyle = new DataGridTextBoxColumn();
            typeColumnStyle.MappingName = "EvtType";
            typeColumnStyle.HeaderText = "Event Type";
            typeColumnStyle.Width = 140;

            /////////////////////////////////////////////////////////////////
            this.beforeCol = new DataColumn("Before", typeof(string));
            this.dataTable.Columns.Add(this.beforeCol);

            DataGridTextBoxColumn beforeColumnStyle = new DataGridTextBoxColumn();
            beforeColumnStyle.MappingName = "Before";
            beforeColumnStyle.HeaderText = "Before";
            beforeColumnStyle.Width = 50;

            /////////////////////////////////////////////////////////////////
            this.successCol = new DataColumn("Success", typeof(string));
            this.dataTable.Columns.Add(this.successCol);

            DataGridTextBoxColumn successColumnStyle = new DataGridTextBoxColumn();
            successColumnStyle.MappingName = "Success";
            successColumnStyle.HeaderText = "Success";
            successColumnStyle.Width = 50;

            /////////////////////////////////////////////////////////////////
            this.formTypeCol = new DataColumn("FormType", typeof(string));
            this.dataTable.Columns.Add(this.formTypeCol);

            DataGridTextBoxColumn formTypeColumnStyle = new DataGridTextBoxColumn();
            formTypeColumnStyle.MappingName = "FormType";
            formTypeColumnStyle.HeaderText = "FormType";
            formTypeColumnStyle.Width = 60;

            /////////////////////////////////////////////////////////////////
            this.formCountCol = new DataColumn("FormCount", typeof(string));
            this.dataTable.Columns.Add(this.formCountCol);

            DataGridTextBoxColumn formCountColumnStyle = new DataGridTextBoxColumn();
            formCountColumnStyle.MappingName = "FormCount";
            formCountColumnStyle.HeaderText = "FormCount";
            formCountColumnStyle.Width = 65;

            /////////////////////////////////////////////////////////////////
            this.itemUIDCol = new DataColumn("ItemUID", typeof(string));
            this.dataTable.Columns.Add(this.itemUIDCol);

            DataGridTextBoxColumn itemUIDColumnStyle = new DataGridTextBoxColumn();
            itemUIDColumnStyle.MappingName = "ItemUID";
            itemUIDColumnStyle.HeaderText = "ItemUID";
            itemUIDColumnStyle.Width = 50;

            /////////////////////////////////////////////////////////////////
            this.colUIDCol = new DataColumn("ColUID", typeof(string));
            this.dataTable.Columns.Add(this.colUIDCol);

            DataGridTextBoxColumn colUIDColumnStyle = new DataGridTextBoxColumn();
            colUIDColumnStyle.MappingName = "ColUID";
            colUIDColumnStyle.HeaderText = "ColUID";
            colUIDColumnStyle.Width = 45;

            /////////////////////////////////////////////////////////////////
            this.rowCol = new DataColumn("Row", typeof(string));
            this.dataTable.Columns.Add(this.rowCol);

            DataGridTextBoxColumn rowColumnStyle = new DataGridTextBoxColumn();
            rowColumnStyle.MappingName = "Row";
            rowColumnStyle.HeaderText = "Row";
            rowColumnStyle.Width = 45;

            /////////////////////////////////////////////////////////////////
            this.formUIDCol = new DataColumn("FormUID", typeof(string));
            this.dataTable.Columns.Add(this.formUIDCol);

            DataGridTextBoxColumn formUIDColumnStyle = new DataGridTextBoxColumn();
            formUIDColumnStyle.MappingName = "FormUID";
            formUIDColumnStyle.HeaderText = "FormUID";
            formUIDColumnStyle.Width = 50;

            /////////////////////////////////////////////////////////////////
            this.formModeCol = new DataColumn("FormMode", typeof(string));
            this.dataTable.Columns.Add(this.formModeCol);

            DataGridTextBoxColumn formModeColumnStyle = new DataGridTextBoxColumn();
            formModeColumnStyle.MappingName = "FormMode";
            formModeColumnStyle.HeaderText = "FormMode";
            formModeColumnStyle.Width = 60;

            /////////////////////////////////////////////////////////////////
            this.charPressedCol = new DataColumn("CharPressed", typeof(string));
            this.dataTable.Columns.Add(this.charPressedCol);

            DataGridTextBoxColumn charPressedColumnStyle = new DataGridTextBoxColumn();
            charPressedColumnStyle.MappingName = "CharPressed";
            charPressedColumnStyle.HeaderText = "CharPressed";
            charPressedColumnStyle.Width = 50;

            /////////////////////////////////////////////////////////////////
            this.innerEventCol = new DataColumn("InnerEvent", typeof(string));
            this.dataTable.Columns.Add(this.innerEventCol);

            DataGridTextBoxColumn innerEventColumnStyle = new DataGridTextBoxColumn();
            innerEventColumnStyle.MappingName = "InnerEvent";
            innerEventColumnStyle.HeaderText = "InnerEvent";
            innerEventColumnStyle.Width = 50;

            /////////////////////////////////////////////////////////////////
            this.itemChangedCol = new DataColumn("ItemChanged", typeof(string));
            this.dataTable.Columns.Add(this.itemChangedCol);

            DataGridTextBoxColumn itemChangedColumnStyle = new DataGridTextBoxColumn();
            itemChangedColumnStyle.MappingName = "ItemChanged";
            itemChangedColumnStyle.HeaderText = "ItemChanged";
            itemChangedColumnStyle.Width = 50;

            /////////////////////////////////////////////////////////////////
            this.modifiersCol = new DataColumn("Modifiers", typeof(string));
            this.dataTable.Columns.Add(this.modifiersCol);

            DataGridTextBoxColumn modifiersColumnStyle = new DataGridTextBoxColumn();
            modifiersColumnStyle.MappingName = "Modifiers";
            modifiersColumnStyle.HeaderText = "Modifiers";
            modifiersColumnStyle.Width = 50;

            /////////////////////////////////////////////////////////////////
            this.popUpIndicatorCol = new DataColumn("PopUpIndicator", typeof(string));
            this.dataTable.Columns.Add(this.popUpIndicatorCol);

            DataGridTextBoxColumn popUpIndicatorColumnStyle = new DataGridTextBoxColumn();
            popUpIndicatorColumnStyle.MappingName = "PopUpIndicator";
            popUpIndicatorColumnStyle.HeaderText = "PopUpIndicator";
            popUpIndicatorColumnStyle.Width = 50;

            /////////////////////////////////////////////////////////////////
            DataGridTableStyle gridStyle = new DataGridTableStyle();
            gridStyle.MappingName = "EventLoggerData";
            gridStyle.GridColumnStyles.Add(nbColumnStyle);
            gridStyle.GridColumnStyles.Add(timeColumnStyle);
            gridStyle.GridColumnStyles.Add(eventColumnStyle);
            gridStyle.GridColumnStyles.Add(typeColumnStyle);
            gridStyle.GridColumnStyles.Add(beforeColumnStyle);
            gridStyle.GridColumnStyles.Add(successColumnStyle);
            gridStyle.GridColumnStyles.Add(formTypeColumnStyle);
            gridStyle.GridColumnStyles.Add(formCountColumnStyle);
            gridStyle.GridColumnStyles.Add(itemUIDColumnStyle);
            gridStyle.GridColumnStyles.Add(colUIDColumnStyle);
            gridStyle.GridColumnStyles.Add(rowColumnStyle);
            gridStyle.GridColumnStyles.Add(formUIDColumnStyle);
            gridStyle.GridColumnStyles.Add(formModeColumnStyle);
            gridStyle.GridColumnStyles.Add(charPressedColumnStyle);
            gridStyle.GridColumnStyles.Add(innerEventColumnStyle);
            gridStyle.GridColumnStyles.Add(itemChangedColumnStyle);
            gridStyle.GridColumnStyles.Add(modifiersColumnStyle);
            gridStyle.GridColumnStyles.Add(popUpIndicatorColumnStyle);

            gridStyle.AlternatingBackColor = Color.LightYellow;

            /////////////////////////////////////////////////////////////////
            this.dataSet = new DataSet();
            this.dataSet.Tables.Add(this.dataTable);

            /////////////////////////////////////////////////////////////////
            this.dataGrid.SetDataBinding(this.dataSet, "EventLoggerData");
            this.dataGrid.TableStyles.Add(gridStyle);
            this.dataGrid.ReadOnly = true;
        }

        /// <summary>
        /// The add row.
        /// </summary>
        /// <param name="eventInfo">
        /// The event info.
        /// </param>
        /// <param name="typeEvent">
        /// The type event.
        /// </param>
        /// <param name="before">
        /// The before.
        /// </param>
        /// <param name="formType">
        /// The form type.
        /// </param>
        /// <param name="formCount">
        /// The form count.
        /// </param>
        protected void AddRow(string eventInfo, string typeEvent, string before, string formType, string formCount)
        {
            this.AddRow(eventInfo, typeEvent, before, formType, formCount, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
        }

        /// <summary>
        /// The add row.
        /// </summary>
        /// <param name="eventInfo">
        /// The event info.
        /// </param>
        /// <param name="typeEvent">
        /// The type event.
        /// </param>
        /// <param name="before">
        /// The before.
        /// </param>
        /// <param name="formType">
        /// The form type.
        /// </param>
        /// <param name="formCount">
        /// The form count.
        /// </param>
        /// <param name="itemUID">
        /// The item uid.
        /// </param>
        /// <param name="colUID">
        /// The col uid.
        /// </param>
        /// <param name="rowNb">
        /// The row nb.
        /// </param>
        /// <param name="formUID">
        /// The form uid.
        /// </param>
        protected void AddRow(string eventInfo, string typeEvent, string before, string formType, string formCount, string itemUID, string colUID, string rowNb, string formUID)
        {
            this.AddRow(eventInfo, typeEvent, before, formType, formCount, itemUID, colUID, rowNb, formUID, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
        }

        /// <summary>
        /// The add row.
        /// </summary>
        /// <param name="eventInfo">
        /// The event info.
        /// </param>
        /// <param name="typeEvent">
        /// The type event.
        /// </param>
        /// <param name="before">
        /// The before.
        /// </param>
        /// <param name="formType">
        /// The form type.
        /// </param>
        /// <param name="formCount">
        /// The form count.
        /// </param>
        /// <param name="itemUID">
        /// The item uid.
        /// </param>
        /// <param name="colUID">
        /// The col uid.
        /// </param>
        /// <param name="rowNb">
        /// The row nb.
        /// </param>
        /// <param name="formUID">
        /// The form uid.
        /// </param>
        /// <param name="formMode">
        /// The form mode.
        /// </param>
        /// <param name="charPressed">
        /// The char pressed.
        /// </param>
        /// <param name="innerEvent">
        /// The inner event.
        /// </param>
        /// <param name="itemChanged">
        /// The item changed.
        /// </param>
        /// <param name="modifiers">
        /// The modifiers.
        /// </param>
        /// <param name="popUpIndicator">
        /// The pop up indicator.
        /// </param>
        /// <param name="success">
        /// The success.
        /// </param>
        protected void AddRow(string eventInfo, string typeEvent, string before, string formType, string formCount, string itemUID, string colUID, string rowNb, string formUID, string formMode, string charPressed, string innerEvent, string itemChanged, string modifiers, string popUpIndicator, string success)
        {
            // should use Control.BeginInvoke: different threads
            this.BeginInvoke(this.evtLog, new object[16] { eventInfo, typeEvent, before, formType, formCount, itemUID, colUID, rowNb, formUID, formMode, charPressed, innerEvent, itemChanged, modifiers, popUpIndicator, success });
        }

        /// <summary>
        /// The add events log.
        /// </summary>
        /// <param name="eventInfo">
        /// The event info.
        /// </param>
        /// <param name="typeEvent">
        /// The type event.
        /// </param>
        /// <param name="before">
        /// The before.
        /// </param>
        /// <param name="formType">
        /// The form type.
        /// </param>
        /// <param name="formCount">
        /// The form count.
        /// </param>
        /// <param name="itemUID">
        /// The item uid.
        /// </param>
        /// <param name="colUID">
        /// The col uid.
        /// </param>
        /// <param name="rowNb">
        /// The row nb.
        /// </param>
        /// <param name="formUID">
        /// The form uid.
        /// </param>
        /// <param name="formMode">
        /// The form mode.
        /// </param>
        /// <param name="charPressed">
        /// The char pressed.
        /// </param>
        /// <param name="innerEvent">
        /// The inner event.
        /// </param>
        /// <param name="itemChanged">
        /// The item changed.
        /// </param>
        /// <param name="modifiers">
        /// The modifiers.
        /// </param>
        /// <param name="popUpIndicator">
        /// The pop up indicator.
        /// </param>
        /// <param name="success">
        /// The success.
        /// </param>
        private void addEventsLog(string eventInfo, string typeEvent, string before, string formType, string formCount, string itemUID, string colUID, string rowNb, string formUID, string formMode, string charPressed, string innerEvent, string itemChanged, string modifiers, string popUpIndicator, string success)
        {
            try
            {
                DataRow row = this.dataTable.NewRow();

                row[this.nbCol] = this.dataTable.Rows.Count;
                row[this.timeCol] = DateTime.Now.TimeOfDay.ToString();

                row[this.eventCol] = eventInfo;
                row[this.typeCol] = typeEvent;
                row[this.beforeCol] = before;
                row[this.successCol] = success;
                row[this.formTypeCol] = formType;
                row[this.formCountCol] = formCount;
                row[this.itemUIDCol] = itemUID;
                row[this.colUIDCol] = colUID;
                row[this.rowCol] = rowNb;
                row[this.formUIDCol] = formUID;
                row[this.formModeCol] = formMode;
                row[this.charPressedCol] = charPressed;
                row[this.innerEventCol] = innerEvent;
                row[this.itemChangedCol] = itemChanged;
                row[this.modifiersCol] = modifiers;
                row[this.popUpIndicatorCol] = popUpIndicator;

                this.dataTable.Rows.Add(row);

                if (!this.clearLog_toolBarButton.Enabled)
                {
                    this.clearLog_toolBarButton.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception thrown " + ex.Message);
            }
        }

        /// <summary>
        /// The tool bar 1_ button click.
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="e">
        /// The e.
        /// </param>
        private void toolBar1_ButtonClick(object sender, ToolBarButtonClickEventArgs e)
        {
            ToolBarButton but = e.Button;

            if (but.ToolTipText == "Connect")
            {
                if (this.Connect())
                {
                    this.connect_toolBarButton.Enabled = false;
                    this.disconnect_toolBarButton.Enabled = true;
                }
                else
                {
                    this.connect_toolBarButton.Enabled = true;
                    this.disconnect_toolBarButton.Enabled = false;
                }
            }
            else if (but.ToolTipText == "Disconnect")
            {
                this.Disconnect();
                this.connect_toolBarButton.Enabled = true;
                this.disconnect_toolBarButton.Enabled = false;
            }
            else if (but.ToolTipText == "Filter Events")
            {
                this.FilterEvents();
            }
            else if (but.ToolTipText == "ClearLog")
            {
                this.dataTable.Rows.Clear();
                this.clearLog_toolBarButton.Enabled = false;
            }
            else if (but.ToolTipText == "Exit")
            {
                this.notifyIcon1.Visible = false;
                Environment.Exit(0);
            }
        }

        /// <summary>
        /// The resize handler.
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="e">
        /// The e.
        /// </param>
        private void ResizeHandler(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.Hide();
            }
        }

        /// <summary>
        /// The close handler.
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="e">
        /// The e.
        /// </param>
        private void CloseHandler(object sender, EventArgs e)
        {
            this.notifyIcon1.Visible = false;
        }

        /// <summary>
        /// The notify icon 1_ mouse double click.
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="e">
        /// The e.
        /// </param>
        private void notifyIcon1_MouseDoubleClick(object sender, EventArgs e)
        {
            if (this.Visible == false)
            {
                this.WindowState = FormWindowState.Maximized;
                this.Show();
                this.Activate();
            }
        }

        #endregion

        #region B1Connection

        /// <summary>
        /// The connect.
        /// </summary>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        public bool Connect()
        {
            try
            {
                this.theGui = new SboGuiApi();
                this.theGui.Connect("0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056");
                this.theAppl = this.theGui.GetApplication(0);

                // connect to the event sink
                this.theAppl = this.theGui.GetApplication(-1);
                this.theAppl.MenuEvent += this.MenuHandler;
                this.theAppl.ItemEvent += this.ItemHandler;
                this.theAppl.AppEvent += this.theAppl_AppEvent;
                this.theAppl.StatusBarEvent += this.StatusBarHandler;
                this.theAppl.ProgressBarEvent += this.ProgressBarHandler;

#if !V2004 // from V2005
                this.theAppl.RightClickEvent += this.RightClickHandler;
                this.theAppl.PrintEvent += this.PrintHandler;
                this.theAppl.ReportDataEvent += this.ReportDataHandler;

#if !V2005 // from V2005_SP01
                this.theAppl.FormDataEvent += this.FormDataHandler;

#if !V2005_SP01 && !V2007 // from 8.8
                this.theAppl.WidgetEvent += this.WidgetEventHandler;
#endif
#endif
#endif
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot connect: " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// The disconnect.
        /// </summary>
        public void Disconnect()
        {
            try
            {
                // disconnect from the event sink
                this.theAppl.MenuEvent -= this.MenuHandler;
                this.theAppl.ItemEvent -= this.ItemHandler;
                this.theAppl.AppEvent -= this.theAppl_AppEvent;
                this.theAppl.StatusBarEvent -= this.StatusBarHandler;
                this.theAppl.ProgressBarEvent -= this.ProgressBarHandler;

#if !V2004 // from V2005
                this.theAppl.RightClickEvent -= this.RightClickHandler;
                this.theAppl.PrintEvent -= this.PrintHandler;
                this.theAppl.ReportDataEvent -= this.ReportDataHandler;
#endif

#if !V2004 && !V2005 // from V2005_SP01
                this.theAppl.FormDataEvent -= this.FormDataHandler;
#endif
            }
            catch (Exception)
            {
            }

            this.theGui = null;
            this.theAppl = null;
        }

        #endregion

        #region EventsManagement

        ////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// The item handler.
        /// </summary>
        /// <param name="FormUId">
        /// The form u id.
        /// </param>
        /// <param name="pVal">
        /// The p val.
        /// </param>
        /// <param name="BubbleEvent">
        /// The bubble event.
        /// </param>
        private void ItemHandler(string FormUId, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST && pVal.BeforeAction == false)
                {
                    IChooseFromListEvent oCFLEvento;
                    oCFLEvento = (IChooseFromListEvent)pVal;
                    string sCFL_ID = oCFLEvento.ChooseFromListUID;
                    string sVal;

                    // exit when nothing had been chosen / found...
                    if (oCFLEvento.SelectedObjects != null && oCFLEvento.SelectedObjects.Rows.Count > 0)
                    {
                        SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                        sVal = oDataTable.GetValue(0, 0).ToString();
                        int i;
                        for (i = 1; i < oDataTable.Rows.Count - 1; i++)
                        {
                            try
                            {
                                sVal = sVal + "; " + oDataTable.GetValue(0, i);
                            }
                            catch (Exception)
                            {
                            }
                        }
                    }
                    else
                    {
                        sVal = "<nothing>";
                    }

                    this.AddRow("Item Event", pVal.EventType.ToString(), pVal.BeforeAction.ToString(), pVal.FormTypeEx, "IDs = " + sVal, "CFL_UID = " + sCFL_ID, pVal.ColUID, pVal.Row.ToString(), pVal.FormUID, pVal.FormMode.ToString(), pVal.CharPressed.ToString(), pVal.InnerEvent.ToString(), pVal.ItemChanged.ToString(), pVal.Modifiers.ToString(), pVal.PopUpIndicator.ToString(), pVal.ActionSuccess.ToString());
                }
                else
                {
                    this.AddRow("Item Event", pVal.EventType.ToString(), pVal.BeforeAction.ToString(), pVal.FormTypeEx, pVal.FormTypeCount.ToString(), pVal.ItemUID, pVal.ColUID, pVal.Row.ToString(), pVal.FormUID, pVal.FormMode.ToString(), pVal.CharPressed.ToString(), pVal.InnerEvent.ToString(), pVal.ItemChanged.ToString(), pVal.Modifiers.ToString(), pVal.PopUpIndicator.ToString(), pVal.ActionSuccess.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception thrown", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// The menu handler.
        /// </summary>
        /// <param name="pVal">
        /// The p val.
        /// </param>
        /// <param name="BubbleEvent">
        /// The bubble event.
        /// </param>
        private void MenuHandler(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                this.AddRow("Menu Event", "<et_MENU_CLICK>", pVal.BeforeAction.ToString(), string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, pVal.MenuUID, string.Empty, string.Empty, pVal.InnerEvent.ToString(), string.Empty, string.Empty, string.Empty, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception thrown", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// The the appl_ app event.
        /// </summary>
        /// <param name="EventType">
        /// The event type.
        /// </param>
        private void theAppl_AppEvent(BoAppEventTypes EventType)
        {
            try
            {
                this.AddRow("Application Event", EventType.ToString(), string.Empty, string.Empty, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception thrown", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// The progress bar handler.
        /// </summary>
        /// <param name="pVal">
        /// The p val.
        /// </param>
        /// <param name="BubbleEvent">
        /// The bubble event.
        /// </param>
        private void ProgressBarHandler(ref ProgressBarEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                this.AddRow("ProgressBar Event", string.Empty, pVal.BeforeAction.ToString(), pVal.EventType.ToString(), string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception thrown", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// The status bar handler.
        /// </summary>
        /// <param name="Text">
        /// The text.
        /// </param>
        /// <param name="MessageType">
        /// The message type.
        /// </param>
        private void StatusBarHandler(string Text, BoStatusBarMessageType MessageType)
        {
            try
            {
                this.AddRow("StatusBar Event", string.Empty, string.Empty, MessageType.ToString(), Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception thrown", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

#if !V2004 // from V2005
        ////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// The right click handler.
        /// </summary>
        /// <param name="pVal">
        /// The p val.
        /// </param>
        /// <param name="BubbleEvent">
        /// The bubble event.
        /// </param>
        private void RightClickHandler(ref ContextMenuInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                this.AddRow("RightClick Event", pVal.EventType.ToString(), pVal.BeforeAction.ToString(), string.Empty, string.Empty, pVal.ItemUID, pVal.ColUID, pVal.Row.ToString(), pVal.FormUID);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception thrown", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// The print handler.
        /// </summary>
        /// <param name="pVal">
        /// The p val.
        /// </param>
        /// <param name="BubbleEvent">
        /// The bubble event.
        /// </param>
        private void PrintHandler(ref PrintEventInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                this.AddRow("Print Event", pVal.EventType.ToString(), pVal.BeforeAction.ToString(), string.Empty, string.Empty, pVal.ItemUID, pVal.ColUID, pVal.Row.ToString(), pVal.FormUID);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception thrown", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// The report data handler.
        /// </summary>
        /// <param name="pVal">
        /// The p val.
        /// </param>
        /// <param name="BubbleEvent">
        /// The bubble event.
        /// </param>
        private void ReportDataHandler(ref ReportDataInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                this.AddRow("ReportData Event", pVal.EventType.ToString(), pVal.BeforeAction.ToString(), string.Empty, string.Empty, pVal.ItemUID, pVal.ColUID, pVal.Row.ToString(), pVal.FormUID);

                if (pVal.BeforeAction)
                {
                    pVal.RegisterForReport(true);
                }
                else
                {
                    pVal.GetPageCount();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception thrown", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

#if !V2004 && !V2005 // from V2005_SP01
        ////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// The form data handler.
        /// </summary>
        /// <param name="pVal">
        /// The p val.
        /// </param>
        /// <param name="BubbleEvent">
        /// The bubble event.
        /// </param>
        private void FormDataHandler(ref BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                this.AddRow("FormData Event", pVal.EventType.ToString(), pVal.BeforeAction.ToString(), pVal.FormTypeEx, "Obj. Type=" + pVal.Type + "; " + pVal.ObjectKey, string.Empty, string.Empty, string.Empty, pVal.FormUID, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, pVal.ActionSuccess.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception thrown", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

#if !V2005_SP01 && !V2007 // from 8.8
        ////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// The widget event handler.
        /// </summary>
        /// <param name="pWidgetData">
        /// The p widget data.
        /// </param>
        /// <param name="BubbleEvent">
        /// The bubble event.
        /// </param>
        private void WidgetEventHandler(ref WidgetData pWidgetData, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                this.AddRow("Widget Event", pWidgetData.EventType.ToString(), string.Empty, (pWidgetData.Form != null) ? pWidgetData.Form.TypeEx : "null", "Position (bottom: " + pWidgetData.Bottom + "; top: " + pWidgetData.Top + ")(left: " + pWidgetData.Left + "; right: " + pWidgetData.Right + ")", pWidgetData.WidgetType + "(" + pWidgetData.WidgetUID + ")", string.Empty, string.Empty, (pWidgetData.Form != null) ? pWidgetData.Form.UniqueID : "null");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception thrown", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

#endif

#endif
#endif

        /// <summary>
        /// The filter events.
        /// </summary>
        private void FilterEvents()
        {
            Filter filter = new Filter();
            try
            {
                if (filter.ShowDialog() == DialogResult.OK)
                {
                    this.ReadFilter();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception thrown", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                filter.Dispose();
            }
        }

        /// <summary>
        /// The read filter.
        /// </summary>
        private void ReadFilter()
        {
            string line = string.Empty;
            string[] infos = new string[2];
            EventFilters evtFilters = new EventFilters();
            EventFilter evtFilter = null;
            bool start = true;
            BoEventTypes lastEvt = BoEventTypes.et_ALL_EVENTS;
            string lastForm;

            // Read and Set Filter
            if (File.Exists(FilterFile))
            {
                StreamReader sr = File.OpenText(FilterFile);

                while ((line = sr.ReadLine()) != null)
                {
                    infos = line.Split(Separator.ToCharArray()[0]);
                    lastForm = infos[1];
                    if (start || lastEvt.ToString() != infos[0])
                    {
                        start = false;
                        lastEvt = (BoEventTypes)Enum.Parse(typeof(BoEventTypes), infos[0]);

                        // Add event to B1 Application
                        evtFilter = evtFilters.Add(lastEvt);
                    }

                    if (infos[1] != string.Empty)
                    {
                        evtFilter.AddEx(infos[1]);
                    }
                }

                this.theAppl.SetFilter(evtFilters);

                sr.Close();
            }
        }

        #endregion

        /// <summary>
        /// The main.
        /// </summary>
        [STAThread]
        private static void Main()
        {
            System.Windows.Forms.Application.Run(new EventLogger());
        }
    }
}