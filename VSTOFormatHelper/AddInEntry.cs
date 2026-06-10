using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace FormatHelper
{
    [ComVisible(true)]
    [Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567890")]
    [ProgId("VSTOFormatHelper.Addin")]
    [ClassInterface(ClassInterfaceType.None)]
    public class AddInEntry : IDTExtensibility2, IRibbonExtensibility
    {
        private Microsoft.Office.Interop.Word.Application _wordApp;
        private static string LogPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WordFormatHelper", "addin.log");

        public AddInEntry()
        {
            Log("AddInEntry constructed");
        }

        private void Log(string msg)
        {
            try
            {
                string dir = Path.GetDirectoryName(LogPath);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                File.AppendAllText(LogPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + msg + "\r\n");
            }
            catch { }
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            Log("OnConnection called");
            _wordApp = (Microsoft.Office.Interop.Word.Application)Application;
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Log("OnDisconnection called");
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom)
        {
            Log("OnStartupComplete called");
        }
        public void OnBeginShutdown(ref Array custom) { }

        public string GetCustomUI(string RibbonID)
        {
            Log("GetCustomUI called with RibbonID: " + RibbonID);
            return RibbonXML;
        }

        private string RibbonXML = @"
<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
  <ribbon>
    <tabs>
      <tab id='FormatTab' label='格式调整'>
        <group id='PresetGroup' label='预设方案'>
          <button id='btnStandard' label='标准报告' size='large' imageMso='AutoFormat' onAction='OnApplyStandard' />
          <button id='btnAudit' label='审计报告' size='large' imageMso='FormatPainter' onAction='OnApplyAudit' />
          <button id='btnCustom' label='自定义格式' size='large' imageMso='CustomPageSetup' onAction='OnApplyCustom' />
          <button id='btnPanel' label='格式设置' size='large' imageMso='PropertySheet' onAction='OnOpenPanel' />
        </group>
        <group id='LevelGroup' label='段落级别'>
          <box id='LevelBox' boxStyle='vertical'>
            <button id='btnH1' label='一级标题' imageMso='Heading1' onAction='OnSetHeading1' />
            <button id='btnH2' label='二级标题' imageMso='Heading2' onAction='OnSetHeading2' />
            <button id='btnH3' label='三级标题' imageMso='Heading3' onAction='OnSetHeading3' />
            <button id='btnH4' label='四级标题' imageMso='Heading4' onAction='OnSetHeading4' />
            <button id='btnNormal' label='正文' imageMso='StyleNormal' onAction='OnSetNormal' />
          </box>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        public void OnApplyStandard(IRibbonControl control)
        {
            ApplyPreset("标准报告");
        }

        public void OnApplyAudit(IRibbonControl control)
        {
            ApplyPreset("审计报告");
        }

        public void OnApplyCustom(IRibbonControl control)
        {
            ApplyPreset("自定义格式");
        }

        public void OnOpenPanel(IRibbonControl control)
        {
            try
            {
                FormatPreset current = PresetManager.GetPresetByName("标准报告");
                FormatPanel panel = new FormatPanel(current);
                if (panel.ShowDialog() == DialogResult.OK)
                {
                    FormatPreset p = panel.GetPreset();
                    ApplyPresetToDoc(p);
                }
            }
            catch (Exception ex)
            {
                Log("OnOpenPanel error: " + ex.Message);
                MessageBox.Show("打开格式面板失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnSetHeading1(IRibbonControl control) { SetHeadingLevel(1); }
        public void OnSetHeading2(IRibbonControl control) { SetHeadingLevel(2); }
        public void OnSetHeading3(IRibbonControl control) { SetHeadingLevel(3); }
        public void OnSetHeading4(IRibbonControl control) { SetHeadingLevel(4); }
        public void OnSetNormal(IRibbonControl control) { SetHeadingLevel(0); }

        private void ApplyPreset(string name)
        {
            try
            {
                FormatPreset p = PresetManager.GetPresetByName(name);
                if (p == null)
                {
                    MessageBox.Show("未找到预设: " + name, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                ApplyPresetToDoc(p);
                MessageBox.Show("已应用预设: " + name, "格式调整", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log("ApplyPreset error: " + ex.Message + "\r\n" + ex.StackTrace);
                MessageBox.Show("应用预设失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplyPresetToDoc(FormatPreset p)
        {
            if (_wordApp == null || _wordApp.ActiveDocument == null)
            {
                MessageBox.Show("请先打开一个Word文档", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            FormatApplier applier = new FormatApplier(_wordApp.ActiveDocument);
            applier.ApplyAll(p);
        }

        private void SetHeadingLevel(int level)
        {
            try
            {
                if (_wordApp == null || _wordApp.ActiveDocument == null)
                {
                    MessageBox.Show("请先打开一个Word文档", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                FormatApplier applier = new FormatApplier(_wordApp.ActiveDocument);
                applier.ApplyHeadingLevel(level);
            }
            catch (Exception ex)
            {
                Log("SetHeadingLevel error: " + ex.Message);
                MessageBox.Show("设置段落级别失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
