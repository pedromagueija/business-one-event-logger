using System;
using System.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;

namespace EventLogger
{
	/// <summary>
	/// Summary description for Filter.
	/// </summary>
	public class Filter : System.Windows.Forms.Form
	{

    private System.Windows.Forms.ComboBox evtType_comboBox;
    private System.Windows.Forms.TextBox formType_textBox;
    private System.Windows.Forms.Label evtType_label;
    private System.Windows.Forms.Label formType_label;
    private System.Windows.Forms.DataGrid dataGrid1;
    private System.Windows.Forms.Button clearall_button;
    private System.Windows.Forms.Button ok_button;
    private System.Windows.Forms.GroupBox groupBox1;
    private System.Windows.Forms.GroupBox groupBox2;
    private System.Windows.Forms.Button addEvt_button;

    private DataTable dataTable;
    private DataColumn nbCol;
    private DataColumn evtTypeCol;
    private DataColumn formTypeCol;
    private DataColumn itemUIDCol;

    private DataSet dataSet;
    private System.Windows.Forms.ContextMenu eventsList_contextMenu;
    private System.Windows.Forms.MenuItem delete_menuItem;
    private System.Windows.Forms.MenuItem deleteall_menuItem;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Filter()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//

      // Fill Event Types Combobox
      FillEventTypes();

      InitializeGrid();

      ReadCurrentFilter();
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
      this.evtType_comboBox = new System.Windows.Forms.ComboBox();
      this.formType_textBox = new System.Windows.Forms.TextBox();
      this.addEvt_button = new System.Windows.Forms.Button();
      this.evtType_label = new System.Windows.Forms.Label();
      this.formType_label = new System.Windows.Forms.Label();
      this.dataGrid1 = new System.Windows.Forms.DataGrid();
      this.eventsList_contextMenu = new System.Windows.Forms.ContextMenu();
      this.delete_menuItem = new System.Windows.Forms.MenuItem();
      this.deleteall_menuItem = new System.Windows.Forms.MenuItem();
      this.clearall_button = new System.Windows.Forms.Button();
      this.ok_button = new System.Windows.Forms.Button();
      this.groupBox1 = new System.Windows.Forms.GroupBox();
      this.groupBox2 = new System.Windows.Forms.GroupBox();
      ((System.ComponentModel.ISupportInitialize)(this.dataGrid1)).BeginInit();
      this.groupBox1.SuspendLayout();
      this.groupBox2.SuspendLayout();
      this.SuspendLayout();
      // 
      // evtType_comboBox
      // 
      this.evtType_comboBox.Location = new System.Drawing.Point(16, 40);
      this.evtType_comboBox.Name = "evtType_comboBox";
      this.evtType_comboBox.Size = new System.Drawing.Size(184, 21);
      this.evtType_comboBox.TabIndex = 0;
      // 
      // formType_textBox
      // 
      this.formType_textBox.Location = new System.Drawing.Point(200, 32);
      this.formType_textBox.Name = "formType_textBox";
      this.formType_textBox.Size = new System.Drawing.Size(96, 20);
      this.formType_textBox.TabIndex = 1;
      this.formType_textBox.Text = "";
      // 
      // addEvt_button
      // 
      this.addEvt_button.Location = new System.Drawing.Point(304, 32);
      this.addEvt_button.Name = "addEvt_button";
      this.addEvt_button.Size = new System.Drawing.Size(96, 23);
      this.addEvt_button.TabIndex = 2;
      this.addEvt_button.Text = "Add Event";
      this.addEvt_button.Click += new System.EventHandler(this.addEvt_button_Click);
      // 
      // evtType_label
      // 
      this.evtType_label.Location = new System.Drawing.Point(16, 24);
      this.evtType_label.Name = "evtType_label";
      this.evtType_label.Size = new System.Drawing.Size(128, 16);
      this.evtType_label.TabIndex = 3;
      this.evtType_label.Text = "EventType";
      // 
      // formType_label
      // 
      this.formType_label.Location = new System.Drawing.Point(208, 24);
      this.formType_label.Name = "formType_label";
      this.formType_label.Size = new System.Drawing.Size(80, 16);
      this.formType_label.TabIndex = 4;
      this.formType_label.Text = "Form Type";
      // 
      // dataGrid1
      // 
      this.dataGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
        | System.Windows.Forms.AnchorStyles.Left) 
        | System.Windows.Forms.AnchorStyles.Right)));
      this.dataGrid1.ContextMenu = this.eventsList_contextMenu;
      this.dataGrid1.DataMember = "";
      this.dataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText;
      this.dataGrid1.Location = new System.Drawing.Point(16, 96);
      this.dataGrid1.Name = "dataGrid1";
      this.dataGrid1.Size = new System.Drawing.Size(392, 168);
      this.dataGrid1.TabIndex = 5;
      // 
      // eventsList_contextMenu
      // 
      this.eventsList_contextMenu.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
                                                                                           this.delete_menuItem,
                                                                                           this.deleteall_menuItem});
      // 
      // delete_menuItem
      // 
      this.delete_menuItem.Index = 0;
      this.delete_menuItem.Text = "Delete current line";
      this.delete_menuItem.Click += new System.EventHandler(this.delete_menuItem_Click);
      // 
      // deleteall_menuItem
      // 
      this.deleteall_menuItem.Index = 1;
      this.deleteall_menuItem.Text = "Delete all";
      this.deleteall_menuItem.Click += new System.EventHandler(this.deleteall_menuItem_Click);
      // 
      // clearall_button
      // 
      this.clearall_button.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
      this.clearall_button.DialogResult = System.Windows.Forms.DialogResult.OK;
      this.clearall_button.Location = new System.Drawing.Point(192, 192);
      this.clearall_button.Name = "clearall_button";
      this.clearall_button.Size = new System.Drawing.Size(88, 24);
      this.clearall_button.TabIndex = 7;
      this.clearall_button.Text = "NO Filter";
      this.clearall_button.Click += new System.EventHandler(this.clearall_button_Click);
      // 
      // ok_button
      // 
      this.ok_button.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
      this.ok_button.DialogResult = System.Windows.Forms.DialogResult.OK;
      this.ok_button.Location = new System.Drawing.Point(296, 192);
      this.ok_button.Name = "ok_button";
      this.ok_button.Size = new System.Drawing.Size(96, 24);
      this.ok_button.TabIndex = 8;
      this.ok_button.Text = "Set Filter";
      this.ok_button.Click += new System.EventHandler(this.ok_button_Click);
      // 
      // groupBox1
      // 
      this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
        | System.Windows.Forms.AnchorStyles.Right)));
      this.groupBox1.Controls.Add(this.formType_textBox);
      this.groupBox1.Controls.Add(this.addEvt_button);
      this.groupBox1.Location = new System.Drawing.Point(8, 8);
      this.groupBox1.Name = "groupBox1";
      this.groupBox1.Size = new System.Drawing.Size(408, 64);
      this.groupBox1.TabIndex = 9;
      this.groupBox1.TabStop = false;
      this.groupBox1.Text = "Add New Event";
      // 
      // groupBox2
      // 
      this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
        | System.Windows.Forms.AnchorStyles.Left) 
        | System.Windows.Forms.AnchorStyles.Right)));
      this.groupBox2.Controls.Add(this.clearall_button);
      this.groupBox2.Controls.Add(this.ok_button);
      this.groupBox2.Location = new System.Drawing.Point(8, 80);
      this.groupBox2.Name = "groupBox2";
      this.groupBox2.Size = new System.Drawing.Size(408, 224);
      this.groupBox2.TabIndex = 10;
      this.groupBox2.TabStop = false;
      this.groupBox2.Text = "Current Filtered Events";
      // 
      // Filter
      // 
      this.AcceptButton = this.addEvt_button;
      this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
      this.ClientSize = new System.Drawing.Size(424, 309);
      this.Controls.Add(this.dataGrid1);
      this.Controls.Add(this.formType_label);
      this.Controls.Add(this.evtType_label);
      this.Controls.Add(this.evtType_comboBox);
      this.Controls.Add(this.groupBox1);
      this.Controls.Add(this.groupBox2);
      this.Name = "Filter";
      this.Text = "Filters Management";
      ((System.ComponentModel.ISupportInitialize)(this.dataGrid1)).EndInit();
      this.groupBox1.ResumeLayout(false);
      this.groupBox2.ResumeLayout(false);
      this.ResumeLayout(false);

    }
		#endregion

    protected void FillEventTypes()
    {
      string[] evtTypes = Enum.GetNames(typeof(SAPbouiCOM.BoEventTypes));
      for (int i = 0; i < evtTypes.Length; i++)
      {
        evtType_comboBox.Items.Add(evtTypes[i]);
      }
    }

    protected void InitializeGrid()
    {
      dataTable = new DataTable("FilterData");
			
      /////////////////////////////////////////////////////////////////

      nbCol = new DataColumn("#", typeof(string));
      dataTable.Columns.Add(nbCol);

      DataGridTextBoxColumn nbColumnStyle = new DataGridTextBoxColumn();
      nbColumnStyle.MappingName = "#";
      nbColumnStyle.HeaderText  = "#";
      nbColumnStyle.Width		= 25;
			
      /////////////////////////////////////////////////////////////////

      evtTypeCol = new DataColumn("EventType", typeof(string));
      dataTable.Columns.Add(evtTypeCol);

      DataGridTextBoxColumn evtTypeColumnStyle = new DataGridTextBoxColumn();
      evtTypeColumnStyle.MappingName = "EventType";
      evtTypeColumnStyle.HeaderText  = "EventType";
      evtTypeColumnStyle.Width		= 200;

      /////////////////////////////////////////////////////////////////

      formTypeCol = new DataColumn("FormType", typeof(string));
      dataTable.Columns.Add(formTypeCol);

      DataGridTextBoxColumn formTypeColumnStyle = new DataGridTextBoxColumn();
      formTypeColumnStyle.MappingName = "FormType";
      formTypeColumnStyle.HeaderText  = "FormType";
      formTypeColumnStyle.Width		= 150;

      /////////////////////////////////////////////////////////////////

      itemUIDCol = new DataColumn("ItemUID", typeof(string));
      dataTable.Columns.Add(itemUIDCol);

      DataGridTextBoxColumn itemUIDColumnStyle = new DataGridTextBoxColumn();
      itemUIDColumnStyle.MappingName = "ItemUID";
      itemUIDColumnStyle.HeaderText  = "ItemUID";
      itemUIDColumnStyle.Width		= 100;

      /////////////////////////////////////////////////////////////////
      DataGridTableStyle gridStyle = new DataGridTableStyle();
      gridStyle.MappingName = "FilterData";
      gridStyle.GridColumnStyles.Add(nbColumnStyle);
      gridStyle.GridColumnStyles.Add(evtTypeColumnStyle);
      gridStyle.GridColumnStyles.Add(formTypeColumnStyle);
      //gridStyle.GridColumnStyles.Add(itemUIDColumnStyle);

      gridStyle.AlternatingBackColor = Color.LightYellow;

      dataSet = new DataSet();
      dataSet.Tables.Add(dataTable);

      dataGrid1.SetDataBinding(dataSet, "FilterData");
      dataGrid1.TableStyles.Add(gridStyle);
      dataGrid1.ReadOnly = true;
    }

    private void addEvt_button_Click(object sender, System.EventArgs e)
    {
      if (evtType_comboBox.SelectedItem == null )
      {
        MessageBox.Show("EventType is a mandatory information.", 
          "Invalid value", MessageBoxButtons.OK, MessageBoxIcon.Error);
        return;
      }
			if (formType_textBox.Text == "")
			{
				MessageBox.Show("Be careful, if FormType is empty all forms will be logged for this event.", 
					"Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}

      int sameEvtRow = -1, evtCount = 0;
      string newEvt = evtType_comboBox.SelectedItem.ToString();
      string formType = formType_textBox.Text;

      // Don't add twice the same line
      foreach (DataRow row in dataTable.Rows)
      {
        if (row[evtTypeCol].ToString().Equals(newEvt))
        {
          sameEvtRow = evtCount;
          if (row[formTypeCol].ToString().Equals(formType))
          {
            MessageBox.Show("Event already assigned to the form","Warning", 
              MessageBoxButtons.OK, MessageBoxIcon.Warning); 
            return;
          }
        }
        evtCount++;
      }

      DataRow newRow = dataTable.NewRow();

      newRow[nbCol] = dataTable.Rows.Count;
      newRow[evtTypeCol] = newEvt;
      newRow[formTypeCol] = formType;

      if (sameEvtRow != -1)
        dataTable.Rows.InsertAt(newRow, sameEvtRow+1);
      else
        dataTable.Rows.Add(newRow);
    }

    private void ok_button_Click(object sender, System.EventArgs e)
    {
      SaveFilteredEvents();

      this.Visible = false;
    }

    private void SaveFilteredEvents()
    {
      if (dataTable.Rows.Count == 0)
        MessageBox.Show("Filter is empty, no event will be shown", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);

      StreamWriter sw = null;
      sw = File.CreateText(EventLogger.FilterFile);

      foreach (DataRow row in dataTable.Rows)
      {
        sw.WriteLine(row[evtTypeCol] +  
          EventLogger.Separator + row[formTypeCol]); 
      }
      sw.Close();
    }

    private void ReadCurrentFilter()
    {
      string line = "";
      string[] infos = new string[2];

      // Read and Set Filter
      if (File.Exists(EventLogger.FilterFile ))
      {
        StreamReader sr = File.OpenText( EventLogger.FilterFile );

        while ((line = sr.ReadLine())!=null) 
        {
          infos = line.Split(EventLogger.Separator.ToCharArray()[0]);

          // Add event dataGrid
          DataRow row = dataTable.NewRow();
          row[nbCol] = dataTable.Rows.Count;
          row[evtTypeCol] = infos[0];
          row[formTypeCol] = infos[1];
          dataTable.Rows.Add(row);
        }
        sr.Close();
      }
    }

    private void delete_menuItem_Click(object sender, System.EventArgs e)
    {
      dataTable.Rows.RemoveAt(dataGrid1.CurrentCell.RowNumber);
    }

    private void deleteall_menuItem_Click(object sender, System.EventArgs e)
    {
      dataTable.Clear();
    }

    private void clearall_button_Click(object sender, System.EventArgs e)
    {
      dataTable.Clear();
      
      // Add et_ALL_EVENTS
      DataRow row = dataTable.NewRow();
      row[nbCol] = dataTable.Rows.Count;
      row[evtTypeCol] = SAPbouiCOM.BoEventTypes.et_ALL_EVENTS;
      row[formTypeCol] = "";
      dataTable.Rows.Add(row);

      SaveFilteredEvents();

      this.Visible = false;
    }


	}
}
