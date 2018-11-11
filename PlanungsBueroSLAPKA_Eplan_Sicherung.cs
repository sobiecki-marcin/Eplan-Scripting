using System;
using System.IO;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Scripting;




public partial class PlanungsBüro_SLAPKA_Eplan_Sicherung : System.Windows.Forms.Form
{
    private System.Windows.Forms.CheckBox CheckBoxProjectCheck;
    private System.Windows.Forms.CheckBox CheckBoxReport;
    private System.Windows.Forms.CheckBox CheckBoxBackup;
    private System.Windows.Forms.SaveFileDialog saveFileDialog1;
    private System.Windows.Forms.TextBox TextBoxLocationPath;
    private System.Windows.Forms.Label label1;
    private System.Windows.Forms.Button ButtonOpenSaveDialog;
    private System.Windows.Forms.CheckBox CheckBoxExportPDF;
    private System.Windows.Forms.RadioButton radioButton1;
    private System.Windows.Forms.RadioButton radioButton2;
    private System.Windows.Forms.Button ButtonOk;
    private System.Windows.Forms.Button ButtonCancel;
    private Label LabelProject;
    private System.Windows.Forms.Label PrueflaufCheck;
    private System.Windows.Forms.Label AuswertenCheck;
    private System.Windows.Forms.Label sichernCheck;
    private System.Windows.Forms.Label PdfCheck;
    string userDefinedDirectory = @"D:\Electric P8\Projekte\";
    public string pathZW = "";
  

   

 

  #region Designer Code

  /// <summary>
  /// Erforderliche Designervariable.
  /// </summary>
  private System.ComponentModel.IContainer components = null;

  /// <summary>
  /// Verwendete Ressourcen bereinigen.
  /// </summary>
  /// <param name="disposing">True, wenn verwaltete Ressourcen
  /// gelöscht werden sollen; andernfalls False.</param>
  protected override void Dispose(bool disposing)
  {
    if (disposing && (components != null))
    {
      components.Dispose();
    }
    base.Dispose(disposing);
  }

  /// <summary>
  /// Erforderliche Methode für die Designerunterstützung.
  /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert
  /// werden.
  /// </summary>
  private void InitializeComponent()
  {
            this.CheckBoxProjectCheck = new System.Windows.Forms.CheckBox();
            this.CheckBoxReport = new System.Windows.Forms.CheckBox();
            this.CheckBoxBackup = new System.Windows.Forms.CheckBox();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.TextBoxLocationPath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.ButtonOpenSaveDialog = new System.Windows.Forms.Button();
            this.CheckBoxExportPDF = new System.Windows.Forms.CheckBox();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.ButtonOk = new System.Windows.Forms.Button();
            this.ButtonCancel = new System.Windows.Forms.Button();
            this.LabelProject = new System.Windows.Forms.Label();
            this.PrueflaufCheck = new System.Windows.Forms.Label();
            this.AuswertenCheck = new System.Windows.Forms.Label();
            this.sichernCheck = new System.Windows.Forms.Label();
            this.PdfCheck = new System.Windows.Forms.Label();
            this.SuspendLayout();




            // 
            // LabelProject
            // 
            this.LabelProject.AutoSize = true;
            this.LabelProject.Location = new System.Drawing.Point(28, 300);
            this.LabelProject.Name = "LabelProject";
            this.LabelProject.Size = new System.Drawing.Size(66, 13);
            this.LabelProject.TabIndex = 5;
            this.LabelProject.Text = "Projektname";

             // 
            // PrueflaufCheck
            // 
            this.PrueflaufCheck.AutoSize = true;
            this.PrueflaufCheck.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.25F);
            this.PrueflaufCheck.ForeColor = System.Drawing.Color.Green;
            this.PrueflaufCheck.Location = new System.Drawing.Point(159, 28);
            this.PrueflaufCheck.Name = "PrueflaufCheck";
            this.PrueflaufCheck.Size = new System.Drawing.Size(17, 17);
            this.PrueflaufCheck.TabIndex = 11;
            //this.PrueflaufCheck.Text = "✓";
            this.PrueflaufCheck.Text = "...in Bearbeitung";
            this.PrueflaufCheck.Visible = false;
            //this.PrueflaufCheck.Click += new System.EventHandler(this.label2_Click);
            // 
            // AuswertenCheck
            // 
            this.AuswertenCheck.AutoSize = true;
            this.AuswertenCheck.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.25F);
            this.AuswertenCheck.ForeColor = System.Drawing.Color.Green;
            this.AuswertenCheck.Location = new System.Drawing.Point(159, 61);
            this.AuswertenCheck.Name = "AuswertenCheck";
            this.AuswertenCheck.Size = new System.Drawing.Size(17, 17);
            this.AuswertenCheck.TabIndex = 12;
            this.AuswertenCheck.Text = "...in Bearbeitung";
            this.AuswertenCheck.Visible = false;
            // 
            // sichernCheck
            // 
            this.sichernCheck.AutoSize = true;
            this.sichernCheck.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.25F);
            this.sichernCheck.ForeColor = System.Drawing.Color.Green;
            this.sichernCheck.Location = new System.Drawing.Point(159, 94);
            this.sichernCheck.Name = "sichernCheck";
            this.sichernCheck.Size = new System.Drawing.Size(17, 17);
            this.sichernCheck.TabIndex = 13;
            this.sichernCheck.Text = "...in Bearbeitung";
            this.sichernCheck.Visible = false;
            // 
            // PdfCheck
            // 
            this.PdfCheck.AutoSize = true;
            this.PdfCheck.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.25F);
            this.PdfCheck.ForeColor = System.Drawing.Color.Green;
            this.PdfCheck.Location = new System.Drawing.Point(159, 203);
            this.PdfCheck.Name = "PdfCheck";
            this.PdfCheck.Size = new System.Drawing.Size(17, 17);
            this.PdfCheck.TabIndex = 14;
            this.PdfCheck.Text = "...in Bearbeitung";
            this.PdfCheck.Visible = false;

            // 
            // CheckBoxProjectCheck
            // 
            this.CheckBoxProjectCheck.AutoSize = true;
            this.CheckBoxProjectCheck.Location = new System.Drawing.Point(28, 30);
            this.CheckBoxProjectCheck.Name = "CheckBoxProjectCheck";
            this.CheckBoxProjectCheck.Size = new System.Drawing.Size(62, 17);
            this.CheckBoxProjectCheck.TabIndex = 0;
            this.CheckBoxProjectCheck.Text = "Prüflauf";
            this.CheckBoxProjectCheck.UseVisualStyleBackColor = true;
            // 
            // CheckBoxReport
            // 
            this.CheckBoxReport.AutoSize = true;
            this.CheckBoxReport.Location = new System.Drawing.Point(28, 63);
            this.CheckBoxReport.Name = "CheckBoxReport";
            this.CheckBoxReport.Size = new System.Drawing.Size(111, 17);
            this.CheckBoxReport.TabIndex = 1;
            this.CheckBoxReport.Text = "Projekt auswerten";
            this.CheckBoxReport.UseVisualStyleBackColor = true;
            // 
            // CheckBoxBackup
            // 
            this.CheckBoxBackup.AutoSize = true;
            this.CheckBoxBackup.Location = new System.Drawing.Point(28, 96);
            this.CheckBoxBackup.Name = "CheckBoxBackup";
            this.CheckBoxBackup.Size = new System.Drawing.Size(96, 17);
            this.CheckBoxBackup.TabIndex = 2;
            this.CheckBoxBackup.Text = "Projekt sichern";
            this.CheckBoxBackup.UseVisualStyleBackColor = true;
            // 
            // TextBoxLocationPath
            // 
            this.TextBoxLocationPath.Location = new System.Drawing.Point(28, 149);
            this.TextBoxLocationPath.Name = "TextBoxLocationPath";
            this.TextBoxLocationPath.Size = new System.Drawing.Size(253, 20);
            this.TextBoxLocationPath.TabIndex = 3;
            this.TextBoxLocationPath.Text = userDefinedDirectory;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(28, 130);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(116, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Sicherungsverzeichnis:";
            // 
            // ButtonOpenSaveDialog
            // 
            this.ButtonOpenSaveDialog.Location = new System.Drawing.Point(287, 149);
            this.ButtonOpenSaveDialog.Name = "ButtonOpenSaveDialog";
            this.ButtonOpenSaveDialog.Size = new System.Drawing.Size(20, 20);
            this.ButtonOpenSaveDialog.TabIndex = 5;
            this.ButtonOpenSaveDialog.Text = "...";
            this.ButtonOpenSaveDialog.UseVisualStyleBackColor = true;
            this.ButtonOpenSaveDialog.Click += new System.EventHandler(this.ButtonOpenSaveDialog_Click);
            // 
            // CheckBoxExportPDF
            // 
            this.CheckBoxExportPDF.AutoSize = true;
            this.CheckBoxExportPDF.Location = new System.Drawing.Point(28, 205);
            this.CheckBoxExportPDF.Name = "CheckBoxExportPDF";
            this.CheckBoxExportPDF.Size = new System.Drawing.Size(93, 17);
            this.CheckBoxExportPDF.TabIndex = 6;
            this.CheckBoxExportPDF.Text = "*.pdf-Ausgabe";
            this.CheckBoxExportPDF.UseVisualStyleBackColor = true;
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Location = new System.Drawing.Point(28, 240);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(51, 17);
            this.radioButton1.TabIndex = 7;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "farbig";
            this.radioButton1.UseVisualStyleBackColor = true;
            this.radioButton1.Checked = true;
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Location = new System.Drawing.Point(134, 240);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(91, 17);
            this.radioButton2.TabIndex = 8;
            this.radioButton2.TabStop = true;
            this.radioButton2.Text = "schwarz/weiß";
            this.radioButton2.UseVisualStyleBackColor = true;
            // 
            // ButtonOk
            // 
            this.ButtonOk.Location = new System.Drawing.Point(162, 339);
            this.ButtonOk.Name = "ButtonOk";
            this.ButtonOk.Size = new System.Drawing.Size(75, 23);
            this.ButtonOk.TabIndex = 9;
            this.ButtonOk.Text = "Start";
            this.ButtonOk.UseVisualStyleBackColor = true;
            this.ButtonOk.Click += new System.EventHandler(this.ButtonOk_Click);
            // 
            // ButtonCancel
            // 
            this.ButtonCancel.Location = new System.Drawing.Point(248, 339);
            this.ButtonCancel.Name = "ButtonCancel";
            this.ButtonCancel.Size = new System.Drawing.Size(75, 23);
            this.ButtonCancel.TabIndex = 10;
            this.ButtonCancel.Text = "Abbrechen";
            this.ButtonCancel.UseVisualStyleBackColor = true;
            this.ButtonCancel.Click += new System.EventHandler(this.ButtonCancel_Click);
            // 
            // PlanungsBüro_SLAPKA_Eplan_Sicherung
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(351, 389);
            this.Controls.Add(this.LabelProject);
            this.Controls.Add(this.ButtonCancel);
            this.Controls.Add(this.ButtonOk);
            this.Controls.Add(this.radioButton2);
            this.Controls.Add(this.radioButton1);
            this.Controls.Add(this.CheckBoxExportPDF);
            this.Controls.Add(this.ButtonOpenSaveDialog);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.TextBoxLocationPath);
            this.Controls.Add(this.CheckBoxBackup);
            this.Controls.Add(this.CheckBoxReport);
            this.Controls.Add(this.CheckBoxProjectCheck);
            this.Controls.Add(this.PdfCheck);
            this.Controls.Add(this.sichernCheck);
            this.Controls.Add(this.AuswertenCheck);
            this.Controls.Add(this.PrueflaufCheck);
            this.Name = "PlanungsBüro_SLAPKA_Eplan_Sicherung";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PlanungsBüro_SLAPKA_Eplan_Sicherung";
            this.Load += new System.EventHandler(this.PlanungsBüro_SLAPKA_Eplan_Sicherung_Load);
            this.ResumeLayout(false);
            this.PerformLayout();
  }

  public PlanungsBüro_SLAPKA_Eplan_Sicherung()
  {
    InitializeComponent();
  }

  #endregion


 [DeclareAction("PlanungsBueroSLAPKA_Eplan_Sicherung")]
  public void Function()
  {
    PlanungsBüro_SLAPKA_Eplan_Sicherung form = new PlanungsBüro_SLAPKA_Eplan_Sicherung();
    form.ShowDialog();

    return;
  }


//Abbrechen Button
  private void ButtonCancel_Click(object sender, System.EventArgs e)
  {
    this.Close();

    return;
  }

//"..." Button
  private void ButtonOpenSaveDialog_Click(object sender, System.EventArgs e)
  {
    string projectName = PathMap.SubstitutePath("$(PROJECTNAME)");
    string date = DateTime.Now.ToString("yyyy.MM.dd");
    string fileName = projectName + "_" + date + ".zw1";

    SaveFileDialog sfd = new SaveFileDialog();
    sfd.DefaultExt = "zw1";
    sfd.FileName = fileName;
    sfd.Filter = "Datensicherung: Projekt (*.zw1)|*.zw1";
    sfd.InitialDirectory = userDefinedDirectory;
    sfd.Title = "Speicherort für Datensicherung wählen:";
    sfd.ValidateNames = true;

    if (sfd.ShowDialog() == DialogResult.OK)
    {
      //FileStream fileStream = File.Create(sfd.FileName);
      //fileStream.Dispose();
      string backupFullFileName = sfd.FileName;
      this.TextBoxLocationPath.Text = backupFullFileName;
    }

    return;
  }


//Start Button
  private void ButtonOk_Click(object sender, System.EventArgs e)
  {
    Cursor.Current = Cursors.WaitCursor;

      while (!string.IsNullOrEmpty(PathMap.SubstitutePath("$(PROJECTNAME)"))) {   
        LabelProject.Text = "Projekt in Bearbeitung: " + PathMap.SubstitutePath("$(PROJECTNAME)");
        //
        this.PrueflaufCheck.Visible = false;
        this.AuswertenCheck.Visible = false;
        this.sichernCheck.Visible = false;
        this.PdfCheck.Visible = false;
        //
        ActionCallingContext ProjektContext = new ActionCallingContext();
        ProjektContext.AddParameter("NOCLOSE", "0");

        CommandLineInterpreter cli = new CommandLineInterpreter();
        ActionCallingContext acc = new ActionCallingContext();


        if (CheckBoxProjectCheck.Checked)
        {
          this.PrueflaufCheck.Visible = true;
          this.PrueflaufCheck.Text = "...in Bearbeitung";
          acc.AddParameter("TYPE", "PROJECT");
          cli.Execute("check", acc);
          this.PrueflaufCheck.Text = "✓";
        }


        if (CheckBoxReport.Checked)
        {
          this.AuswertenCheck.Visible = true;
          this.AuswertenCheck.Text = "...in Bearbeitung";
          cli.Execute("reports");
          this.AuswertenCheck.Text = "✓";
          
        }


        if (CheckBoxBackup.Checked)
        {
            this.sichernCheck.Visible = true;
            this.sichernCheck.Text = "...in Bearbeitung";


            string projectName = PathMap.SubstitutePath("$(PROJECTNAME)");
            string date = DateTime.Now.ToString("yyyy.MM.dd");
            string backupFileName = projectName + "_" + date + ".zw1";
            string backupPath = this.TextBoxLocationPath.Text;
            string pattern = @"[^\\]*[0-9-A-Z-a-z]_([^;]*)zw1";
            // [^\\]*[0-9-A-Z-a-z]_([^;]*)zw1
            // [^\\]*[\w]_([^;]*)zw1
            string pathZW = Regex.Replace(backupPath, pattern, "");

      
            
            acc.AddParameter("BACKUPAMOUNT", "BACKUPAMOUNT_ALL");
            acc.AddParameter("BACKUPMEDIA", "DISK");
            acc.AddParameter("BACKUPMETHOD", "BACKUP");
            acc.AddParameter("COMPRESSPRJ", "1");
            acc.AddParameter("INCLEXTDOCS", "1");
            acc.AddParameter("INCLIMAGES", "1");
            acc.AddParameter("ARCHIVENAME", backupFileName);
            acc.AddParameter("DESTINATIONPATH", pathZW);
            acc.AddParameter("TYPE", "PROJECT");

            cli.Execute("backup", acc);
            this.sichernCheck.Text = "✓";
        }


        if (CheckBoxExportPDF.Checked)
        {
            this.PdfCheck.Visible = true;
            this.PdfCheck.Text = "...in Bearbeitung";

            string projectName = PathMap.SubstitutePath("$(PROJECTNAME)");
            string date = DateTime.Now.ToString("yyyy.MM.dd");
            string backupFileNamePDF = projectName + "_" + date + ".pdf";
            string backupPath = this.TextBoxLocationPath.Text;
            string pattern = @"[^\\]*[0-9-A-Z-a-z]_([^;]*)zw1";
            // [^\\]*[0-9-A-Z-a-z]_([^;]*)zw1
            // [^\\]*[\w]_([^;]*)zw1
            string pathZW = Regex.Replace(backupPath, pattern, "");
            string backupFullFileNamePDF = Path.Combine(pathZW, backupFileNamePDF);

          

            //MessageBox.Show(backupFullFileNamePDF);
            
            acc.AddParameter("TYPE", "PDFPROJECTSCHEME");
            if(radioButton2.Checked){
              this.radioButton1.Checked = false;
              acc.AddParameter("BLACKWHITE", "1");
            } else if (radioButton1.Checked) {
              acc.AddParameter("BLACKWHITE", "0");
              this.radioButton2.Checked = false;
            } else {
              acc.AddParameter("BLACKWHITE", "0");
            }
            acc.AddParameter("EXPORTFILE", backupFullFileNamePDF);
            acc.AddParameter("EXPORTSCHEME", "EPLAN_default_value");

            cli.Execute("export", acc);
            this.PdfCheck.Text = "✓";
        }

            

            new CommandLineInterpreter().Execute("XPrjActionProjectClose", ProjektContext);

           
 }
    Cursor.Current = Cursors.Default;  
    this.Close();

    return;
  }
  
    private void PlanungsBüro_SLAPKA_Eplan_Sicherung_Load(object sender, System.EventArgs e)
  {
    LabelProject.Text = "Projekt in Bearbeitung: " + PathMap.SubstitutePath("$(PROJECTNAME)");
  }

}