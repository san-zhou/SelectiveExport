using System;
using System.Windows.Forms;
using KeePass.Plugins;
using KeePass.Forms;
using KeePassLib;
using System.Collections.Generic;
using KeePassLib.Security;
using System.IO;
using System.Xml;
using System.Linq;

namespace SelectiveExport
{
  public sealed class SelectiveExportExt : Plugin
  {
    private IPluginHost m_host = null;
    private ToolStripMenuItem m_menuItem = null;

    public override bool Initialize(IPluginHost host)
    {
      if (host == null) return false;
      m_host = host;

      // 创建菜单项
      m_menuItem = new ToolStripMenuItem();
      m_menuItem.Text = "导出选定项...";
      m_menuItem.Click += OnMenuExport;

      // 添加到工具菜单
      m_host.MainWindow.ToolsMenu.DropDownItems.Add(m_menuItem);

      return true;
    }

    private void OnMenuExport(object sender, EventArgs e)
    {
      var mainForm = m_host.MainWindow;
      var selectedEntries = mainForm.GetSelectedEntries();
      var selectedGroups = new List<PwGroup>();

      // 获取当前选中的组
      if (mainForm.ActiveDatabase != null && mainForm.ActiveDatabase.RootGroup != null)
      {
        var currentGroup = mainForm.GetSelectedGroup();
        if (currentGroup != null)
        {
          selectedGroups.Add(currentGroup);
        }
      }

      if ((selectedEntries == null || selectedEntries.Count() == 0) &&
          (selectedGroups == null || selectedGroups.Count == 0))
      {
        MessageBox.Show("请先选择要导出的条目或分组！", "提示",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
        return;
      }

      using (SaveFileDialog sfd = new SaveFileDialog())
      {
        sfd.Filter = "XML 文件 (*.xml)|*.xml|所有文件 (*.*)|*.*";
        sfd.DefaultExt = "xml";
        sfd.FileName = "KeePassExport.xml";

        if (sfd.ShowDialog() == DialogResult.OK)
        {
          ExportData(sfd.FileName, selectedEntries, selectedGroups);
        }
      }
    }

    private void ExportData(string fileName,
        IEnumerable<PwEntry> entries,
        IEnumerable<PwGroup> groups)
    {
      using (XmlWriter writer = XmlWriter.Create(fileName,
          new XmlWriterSettings { Indent = true }))
      {
        writer.WriteStartDocument();
        writer.WriteStartElement("KeePassExport");

        // 导出选中的条目
        if (entries != null)
        {
          writer.WriteStartElement("Entries");
          foreach (var entry in entries)
          {
            WriteEntry(writer, entry);
          }
          writer.WriteEndElement();
        }

        // 导出选中的分组
        if (groups != null)
        {
          writer.WriteStartElement("Groups");
          foreach (var group in groups)
          {
            WriteGroup(writer, group);
          }
          writer.WriteEndElement();
        }

        writer.WriteEndElement();
        writer.WriteEndDocument();
      }

      MessageBox.Show("导出完成！", "成功",
          MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void WriteEntry(XmlWriter writer, PwEntry entry)
    {
      writer.WriteStartElement("Entry");
      writer.WriteElementString("Title", entry.Strings.ReadSafe("Title"));
      writer.WriteElementString("UserName", entry.Strings.ReadSafe("UserName"));
      // writer.WriteElementString("Password", entry.Strings.ReadSafe("Password"));
      // 使用 CDATA 包装密码字段，避免特殊字符被转义
      writer.WriteStartElement("Password");
      writer.WriteCData(entry.Strings.ReadSafe("Password"));
      writer.WriteEndElement();

      writer.WriteElementString("URL", entry.Strings.ReadSafe("URL"));
      writer.WriteElementString("Notes", entry.Strings.ReadSafe("Notes"));
      writer.WriteEndElement();
    }

    private void WriteGroup(XmlWriter writer, PwGroup group)
    {
      writer.WriteStartElement("Group");
      writer.WriteElementString("Name", group.Name);

      // 递归导出组内的条目
      foreach (var entry in group.Entries)
      {
        WriteEntry(writer, entry);
      }

      // 递归导出子组
      foreach (var subGroup in group.Groups)
      {
        WriteGroup(writer, subGroup);
      }

      writer.WriteEndElement();
    }

    public override void Terminate()
    {
      if (m_menuItem != null)
      {
        m_host.MainWindow.ToolsMenu.DropDownItems.Remove(m_menuItem);
      }
    }
  }
}