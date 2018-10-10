#----------------------------------------------
# Generated Form Function
#----------------------------------------------
function Show-GUI_psf {

	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	#endregion Import Assemblies

	#----------------------------------------------
	#region Define SAPIEN Types
	#----------------------------------------------
	try{
		[FolderBrowserModernDialog] | Out-Null
	}
	catch
	{
		Add-Type -ReferencedAssemblies ('System.Windows.Forms') -TypeDefinition  @" 
		using System;
		using System.Windows.Forms;
		using System.Reflection;

        namespace SAPIENTypes
        {
		    public class FolderBrowserModernDialog : System.Windows.Forms.CommonDialog
            {
                private System.Windows.Forms.OpenFileDialog fileDialog;
                public FolderBrowserModernDialog()
                {
                    fileDialog = new System.Windows.Forms.OpenFileDialog();
                    fileDialog.Filter = "Folders|\n";
                    fileDialog.AddExtension = false;
                    fileDialog.CheckFileExists = false;
                    fileDialog.DereferenceLinks = true;
                    fileDialog.Multiselect = false;
                    fileDialog.Title = "Select a folder";
                }

                public string Title
                {
                    get { return fileDialog.Title; }
                    set { fileDialog.Title = value; }
                }

                public string InitialDirectory
                {
                    get { return fileDialog.InitialDirectory; }
                    set { fileDialog.InitialDirectory = value; }
                }
                
                public string SelectedPath
                {
                    get { return fileDialog.FileName; }
                    set { fileDialog.FileName = value; }
                }

                object InvokeMethod(Type type, object obj, string method, object[] parameters)
                {
                    MethodInfo methInfo = type.GetMethod(method, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                    return methInfo.Invoke(obj, parameters);
                }

                bool ShowOriginalBrowserDialog(IntPtr hwndOwner)
                {
                    using(FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
                    {
                        folderBrowserDialog.Description = this.Title;
                        folderBrowserDialog.SelectedPath = !string.IsNullOrEmpty(this.SelectedPath) ? this.SelectedPath : this.InitialDirectory;
                        folderBrowserDialog.ShowNewFolderButton = false;
                        if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                        {
                            fileDialog.FileName = folderBrowserDialog.SelectedPath;
                            return true;
                        }
                        return false;
                    }
                }

                protected override bool RunDialog(IntPtr hwndOwner)
                {
                    if (Environment.OSVersion.Version.Major >= 6)
                    {      
                        try
                        {
                            bool flag = false;
                            System.Reflection.Assembly assembly = Assembly.Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089");
                            Type typeIFileDialog = assembly.GetType("System.Windows.Forms.FileDialogNative").GetNestedType("IFileDialog", BindingFlags.NonPublic);
                            uint num = 0;
                            object dialog = InvokeMethod(fileDialog.GetType(), fileDialog, "CreateVistaDialog", null);
                            InvokeMethod(fileDialog.GetType(), fileDialog, "OnBeforeVistaDialog", new object[] { dialog });
                            uint options = (uint)InvokeMethod(typeof(System.Windows.Forms.FileDialog), fileDialog, "GetOptions", null) | (uint)0x20;
                            InvokeMethod(typeIFileDialog, dialog, "SetOptions", new object[] { options });
                            Type vistaDialogEventsType = assembly.GetType("System.Windows.Forms.FileDialog").GetNestedType("VistaDialogEvents", BindingFlags.NonPublic);
                            object pfde = Activator.CreateInstance(vistaDialogEventsType, fileDialog);
                            object[] parameters = new object[] { pfde, num };
                            InvokeMethod(typeIFileDialog, dialog, "Advise", parameters);
                            num = (uint)parameters[1];
                            try
                            {
                                int num2 = (int)InvokeMethod(typeIFileDialog, dialog, "Show", new object[] { hwndOwner });
                                flag = 0 == num2;
                            }
                            finally
                            {
                                InvokeMethod(typeIFileDialog, dialog, "Unadvise", new object[] { num });
                                GC.KeepAlive(pfde);
                            }
                            return flag;
                        }
                        catch
                        {
                            return ShowOriginalBrowserDialog(hwndOwner);
                        }
                    }
                    else
                        return ShowOriginalBrowserDialog(hwndOwner);
                }

                public override void Reset()
                {
                    fileDialog.Reset();
                }
            }
       }
"@ -IgnoreWarnings | Out-Null
	}
	#endregion Define SAPIEN Types

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$form1 = New-Object 'System.Windows.Forms.Form'
	$buttonGenerujRaport = New-Object 'System.Windows.Forms.Button'
	$labelProszęKatalogDocelow = New-Object 'System.Windows.Forms.Label'
	$labelKatalogDocelowy = New-Object 'System.Windows.Forms.Label'
	$buttonWybierzKatalogDocelo = New-Object 'System.Windows.Forms.Button'
	$labelProszęWybraćPlikŹród = New-Object 'System.Windows.Forms.Label'
	$buttonWybierzPlikRaportu = New-Object 'System.Windows.Forms.Button'
	$labelPlikRaportuSAP = New-Object 'System.Windows.Forms.Label'
	$openfiledialog1 = New-Object 'System.Windows.Forms.OpenFileDialog'
	$folderbrowsermoderndialog1 = New-Object 'SAPIENTypes.FolderBrowserModernDialog'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Generated Form Objects

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	$form1.SuspendLayout()
	#
	# form1
	#
	$form1.Controls.Add($buttonGenerujRaport)
	$form1.Controls.Add($labelProszęKatalogDocelow)
	$form1.Controls.Add($labelKatalogDocelowy)
	$form1.Controls.Add($buttonWybierzKatalogDocelo)
	$form1.Controls.Add($labelProszęWybraćPlikŹród)
	$form1.Controls.Add($buttonWybierzPlikRaportu)
	$form1.Controls.Add($labelPlikRaportuSAP)
	$form1.AutoScaleDimensions = '10, 20'
	$form1.AutoScaleMode = 'Font'
	$form1.ClientSize = '760, 408'
	$form1.Name = 'form1'
	$form1.Text = 'Form'
	$form1.add_Load($form1_Load)
	#
	# buttonGenerujRaport
	#
	$buttonGenerujRaport.Font = 'Microsoft Sans Serif, 8pt'
	$buttonGenerujRaport.Location = '494, 129'
	$buttonGenerujRaport.Margin = '5, 5, 5, 5'
	$buttonGenerujRaport.Name = 'buttonGenerujRaport'
	$buttonGenerujRaport.Size = '230, 33'
	$buttonGenerujRaport.TabIndex = 6
	$buttonGenerujRaport.Text = 'Generuj raport'
	$buttonGenerujRaport.UseCompatibleTextRendering = $True
	$buttonGenerujRaport.UseVisualStyleBackColor = $True
	$buttonGenerujRaport.add_Click($buttonGenerujRaport_Click)
	#
	# labelProszęKatalogDocelow
	#
	$labelProszęKatalogDocelow.AutoSize = $True
	$labelProszęKatalogDocelow.Enabled = $False
	$labelProszęKatalogDocelow.Location = '179, 80'
	$labelProszęKatalogDocelow.Margin = '5, 0, 5, 0'
	$labelProszęKatalogDocelow.Name = 'labelProszęKatalogDocelow'
	$labelProszęKatalogDocelow.Size = '192, 24'
	$labelProszęKatalogDocelow.TabIndex = 5
	$labelProszęKatalogDocelow.Text = 'proszę katalog docelowy'
	$labelProszęKatalogDocelow.UseCompatibleTextRendering = $True
	$labelProszęKatalogDocelow.add_Click($labelProszęKatalogDocelow_Click)
	#
	# labelKatalogDocelowy
	#
	$labelKatalogDocelowy.AutoSize = $True
	$labelKatalogDocelowy.Location = '25, 80'
	$labelKatalogDocelowy.Margin = '5, 0, 5, 0'
	$labelKatalogDocelowy.Name = 'labelKatalogDocelowy'
	$labelKatalogDocelowy.Size = '144, 24'
	$labelKatalogDocelowy.TabIndex = 4
	$labelKatalogDocelowy.Text = 'Katalog docelowy:'
	$labelKatalogDocelowy.UseCompatibleTextRendering = $True
	$labelKatalogDocelowy.add_Click($labelKatalogDocelowy_Click)
	#
	# buttonWybierzKatalogDocelo
	#
	$buttonWybierzKatalogDocelo.Font = 'Microsoft Sans Serif, 8pt'
	$buttonWybierzKatalogDocelo.Location = '494, 74'
	$buttonWybierzKatalogDocelo.Margin = '5, 5, 5, 5'
	$buttonWybierzKatalogDocelo.Name = 'buttonWybierzKatalogDocelo'
	$buttonWybierzKatalogDocelo.Size = '230, 33'
	$buttonWybierzKatalogDocelo.TabIndex = 3
	$buttonWybierzKatalogDocelo.Text = 'Wybierz katalog docelowy'
	$buttonWybierzKatalogDocelo.UseCompatibleTextRendering = $True
	$buttonWybierzKatalogDocelo.UseVisualStyleBackColor = $True
	$buttonWybierzKatalogDocelo.add_Click($buttonWybierzKatalogDocelo_Click)
	#
	# labelProszęWybraćPlikŹród
	#
	$labelProszęWybraćPlikŹród.AutoSize = $True
	$labelProszęWybraćPlikŹród.Enabled = $False
	$labelProszęWybraćPlikŹród.Location = '179, 32'
	$labelProszęWybraćPlikŹród.Margin = '5, 0, 5, 0'
	$labelProszęWybraćPlikŹród.Name = 'labelProszęWybraćPlikŹród'
	$labelProszęWybraćPlikŹród.Size = '217, 24'
	$labelProszęWybraćPlikŹród.TabIndex = 2
	$labelProszęWybraćPlikŹród.Text = 'proszę wybrać plik źródłowy'
	$labelProszęWybraćPlikŹród.UseCompatibleTextRendering = $True
	$labelProszęWybraćPlikŹród.add_Click($labelProszęWybraćPlikŹród_Click)
	#
	# buttonWybierzPlikRaportu
	#
	$buttonWybierzPlikRaportu.Font = 'Microsoft Sans Serif, 8pt'
	$buttonWybierzPlikRaportu.Location = '494, 32'
	$buttonWybierzPlikRaportu.Margin = '5, 5, 5, 5'
	$buttonWybierzPlikRaportu.Name = 'buttonWybierzPlikRaportu'
	$buttonWybierzPlikRaportu.Size = '230, 32'
	$buttonWybierzPlikRaportu.TabIndex = 1
	$buttonWybierzPlikRaportu.Text = 'Wybierz plik raportu'
	$buttonWybierzPlikRaportu.UseCompatibleTextRendering = $True
	$buttonWybierzPlikRaportu.UseVisualStyleBackColor = $True
	$buttonWybierzPlikRaportu.add_Click($buttonWybierzPlikRaportu_Click)
	#
	# labelPlikRaportuSAP
	#
	$labelPlikRaportuSAP.AutoSize = $True
	$labelPlikRaportuSAP.Location = '25, 32'
	$labelPlikRaportuSAP.Margin = '5, 0, 5, 0'
	$labelPlikRaportuSAP.Name = 'labelPlikRaportuSAP'
	$labelPlikRaportuSAP.Size = '146, 24'
	$labelPlikRaportuSAP.TabIndex = 0
	$labelPlikRaportuSAP.Text = 'Plik raportu (SAP):'
	$labelPlikRaportuSAP.UseCompatibleTextRendering = $True
	$labelPlikRaportuSAP.add_Click($labelPlikRaportuSAP_Click)
	#
	# openfiledialog1
	#
	$openfiledialog1.FileName = 'openfiledialog1'
	$openfiledialog1.add_FileOk($openfiledialog1_FileOk)
	#
	# folderbrowsermoderndialog1
	#
	$folderbrowsermoderndialog1.InitialDirectory = "$env:USERPROFILE\Desktop"
	$form1.ResumeLayout()
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $form1.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$form1.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$form1.add_FormClosed($Form_Cleanup_FormClosed)
	#Show the Form
	return $form1.ShowDialog()

} #End Function

#Call the form
Show-GUI_psf | Out-Null
