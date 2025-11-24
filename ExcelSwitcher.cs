/*/ role exeProgram; /*/
using Au;
using Au.Types;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;

// 設定ファイルのパス
string settingsPath = folders.ThisAppDocuments + "ExcelSwitcherSettings.json";

// Excelスイッチャーのメインフォーム
var form = new ExcelSwitcherForm(settingsPath);
form.ShowDialog();

// グループ設定を保存するクラス
public class GroupSettings
{
	public List<GroupData> Groups { get; set; } = new List<GroupData>();
}

public class GroupData
{
	public string Name { get; set; }
	public List<string> Files { get; set; } = new List<string>();
}

// メインフォームクラス
public class ExcelSwitcherForm : Form
{
	private TreeView treeView;
	private string settingsPath;
	private GroupSettings settings;
	private System.Windows.Forms.Timer refreshTimer;
	private HashSet<IntPtr> lastExcelWindows = new HashSet<IntPtr>();

	public ExcelSwitcherForm(string settingsPath)
	{
		this.settingsPath = settingsPath;
		LoadSettings();
		InitializeUI();
		LoadExcelWindows();
		StartRefreshTimer();
	}

	private void InitializeUI()
	{
		// フォーム設定
		this.Text = "Excel Switcher";
		this.Size = new Size(500, 600);
		this.StartPosition = FormStartPosition.CenterScreen;
		this.KeyPreview = true;
		this.TopMost = true;

		// TreeView設定
		treeView = new TreeView
		{
			Dock = DockStyle.Fill,
			Font = new Font("Yu Gothic UI", 11),
			AllowDrop = true,
			HideSelection = false
		};

		// イベントハンドラ
		treeView.NodeMouseDoubleClick += (s, e) => ActivateExcel(e.Node);
		treeView.ItemDrag += TreeView_ItemDrag;
		treeView.DragEnter += TreeView_DragEnter;
		treeView.DragOver += TreeView_DragOver;
		treeView.DragDrop += TreeView_DragDrop;
		treeView.MouseClick += TreeView_MouseClick;

		this.KeyDown += (s, e) =>
		{
			if (e.KeyCode == Keys.Escape) this.Close();
			if (e.KeyCode == Keys.Enter && treeView.SelectedNode != null)
			{
				ActivateExcel(treeView.SelectedNode);
			}
		};

		this.FormClosing += (s, e) => SaveSettings();
		this.Controls.Add(treeView);
	}

	private void LoadSettings()
	{
		if (File.Exists(settingsPath))
		{
			try
			{
				string json = File.ReadAllText(settingsPath);
				settings = JsonSerializer.Deserialize<GroupSettings>(json) ?? new GroupSettings();
			}
			catch
			{
				settings = new GroupSettings();
			}
		}
		else
		{
			settings = new GroupSettings();
		}

		if (!settings.Groups.Any(g => g.Name == "未分類"))
		{
			settings.Groups.Insert(0, new GroupData { Name = "未分類" });
		}
	}

	private void SaveSettings()
	{
		try
		{
			settings.Groups.Clear();
			foreach (TreeNode groupNode in treeView.Nodes)
			{
				var group = new GroupData { Name = groupNode.Text };
				foreach (TreeNode fileNode in groupNode.Nodes)
				{
					group.Files.Add(fileNode.Text);
				}
				settings.Groups.Add(group);
			}

			string json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
			File.WriteAllText(settingsPath, json);
		}
		catch (Exception ex)
		{
			print.it($"設定の保存に失敗: {ex.Message}");
		}
	}

	private void LoadExcelWindows()
	{
		treeView.Nodes.Clear();
		var currentExcelWindows = new HashSet<IntPtr>();

		var excelWindows = wnd.findAll(cn: "XLMAIN");
		var excelData = new List<(wnd w, string title)>();

		foreach (var w in excelWindows)
		{
			if (w.IsVisible)
			{
				string title = w.Name;
				if (!string.IsNullOrWhiteSpace(title))
				{
					excelData.Add((w, title));
					currentExcelWindows.Add(w.Handle);
				}
			}
		}

		foreach (var group in settings.Groups)
		{
			var groupNode = new TreeNode(group.Name) { Tag = null };
			groupNode.NodeFont = new Font(treeView.Font, FontStyle.Bold);
			groupNode.ForeColor = Color.DarkBlue;

			foreach (var fileName in group.Files)
			{
				var matchedExcel = excelData.FirstOrDefault(e => e.title.Contains(fileName));
				if (matchedExcel.w.Is0 == false)
				{
					var fileNode = new TreeNode(matchedExcel.title) { Tag = matchedExcel.w.Handle };
					groupNode.Nodes.Add(fileNode);
					excelData.Remove(matchedExcel);
				}
			}

			treeView.Nodes.Add(groupNode);
			groupNode.Expand();
		}

		if (excelData.Count > 0)
		{
			var uncategorizedNode = treeView.Nodes.Cast<TreeNode>().FirstOrDefault(n => n.Text == "未分類");
			if (uncategorizedNode == null)
			{
				uncategorizedNode = new TreeNode("未分類") { Tag = null };
				uncategorizedNode.NodeFont = new Font(treeView.Font, FontStyle.Bold);
				uncategorizedNode.ForeColor = Color.DarkBlue;
				treeView.Nodes.Insert(0, uncategorizedNode);
			}

			foreach (var excel in excelData)
			{
				var fileNode = new TreeNode(excel.title) { Tag = excel.w.Handle };
				uncategorizedNode.Nodes.Add(fileNode);
			}

			uncategorizedNode.Expand();
		}

		lastExcelWindows = currentExcelWindows;
	}

	private void StartRefreshTimer()
	{
		refreshTimer = new System.Windows.Forms.Timer { Interval = 1000 };
		refreshTimer.Tick += (s, e) => CheckAndRefreshExcelWindows();
		refreshTimer.Start();
	}

	private void CheckAndRefreshExcelWindows()
	{
		var currentExcelWindows = new HashSet<IntPtr>();
		var excelWindows = wnd.findAll(cn: "XLMAIN");

		foreach (var w in excelWindows)
		{
			if (w.IsVisible)
			{
				currentExcelWindows.Add(w.Handle);
			}
		}

		if (!currentExcelWindows.SetEquals(lastExcelWindows))
		{
			LoadExcelWindows();
		}
	}

	private void ActivateExcel(TreeNode node)
	{
		if (node.Tag == null) return;

		wnd w = (wnd)(IntPtr)node.Tag;

		if (!w.IsAlive)
		{
			dialog.show("選択したExcelウィンドウは既に閉じられています。", "エラー");
			return;
		}

		if (w.IsMinimized)
		{
			w.ShowNotMinMax();
		}

		w.Activate();
	}

	private void TreeView_ItemDrag(object sender, ItemDragEventArgs e)
	{
		if (e.Item is TreeNode node)
		{
			DoDragDrop(node, DragDropEffects.Move);
		}
	}

	private void TreeView_DragEnter(object sender, DragEventArgs e)
	{
		e.Effect = DragDropEffects.Move;
	}

	private void TreeView_DragOver(object sender, DragEventArgs e)
	{
		Point targetPoint = treeView.PointToClient(new Point(e.X, e.Y));
		treeView.SelectedNode = treeView.GetNodeAt(targetPoint);
		e.Effect = DragDropEffects.Move;
	}

	private void TreeView_DragDrop(object sender, DragEventArgs e)
	{
		if (!(e.Data.GetData(typeof(TreeNode)) is TreeNode draggedNode)) return;

		Point targetPoint = treeView.PointToClient(new Point(e.X, e.Y));
		TreeNode targetNode = treeView.GetNodeAt(targetPoint);

		if (draggedNode.Tag == null)
		{
			int newIndex = targetNode == null ? treeView.Nodes.Count : targetNode.Index;
			draggedNode.Remove();
			treeView.Nodes.Insert(newIndex, draggedNode);
			SaveSettings();
		}
		else
		{
			if (targetNode != null)
			{
				TreeNode targetGroup = targetNode.Tag == null ? targetNode : targetNode.Parent;

				if (targetGroup != null)
				{
					draggedNode.Remove();
					targetGroup.Nodes.Add(draggedNode);
					targetGroup.Expand();
					SaveSettings();
				}
			}
		}
	}

	private void TreeView_MouseClick(object sender, MouseEventArgs e)
	{
		if (e.Button == MouseButtons.Right)
		{
			TreeNode clickedNode = treeView.GetNodeAt(e.Location);
			var menu = new ContextMenuStrip();

			// ✅ 修正: グループノードをクリックした場合
			if (clickedNode != null && clickedNode.Tag == null)
			{
				// グループ削除
				menu.Items.Add("グループを削除", null, (s, ev) =>
				{
					if (clickedNode.Text == "未分類")
					{
						dialog.show("「未分類」グループは削除できません。", "エラー");
						return;
					}

					var result = dialog.show("グループを削除しますか？\n（グループ内のファイルは「未分類」に移動されます）",
						"確認", "はい|いいえ", icon: DIcon.Warning);
					
					if (result == 1)
					{
						var uncategorizedNode = treeView.Nodes.Cast<TreeNode>().FirstOrDefault(n => n.Text == "未分類");
						if (uncategorizedNode == null)
						{
							uncategorizedNode = new TreeNode("未分類") { Tag = null };
							uncategorizedNode.NodeFont = new Font(treeView.Font, FontStyle.Bold);
							uncategorizedNode.ForeColor = Color.DarkBlue;
							treeView.Nodes.Insert(0, uncategorizedNode);
						}

						while (clickedNode.Nodes.Count > 0)
						{
							var fileNode = clickedNode.Nodes[0];
							fileNode.Remove();
							uncategorizedNode.Nodes.Add(fileNode);
						}

						clickedNode.Remove();
						uncategorizedNode.Expand();
						SaveSettings();
					}
				});

				// グループ名変更
				menu.Items.Add("グループ名を変更", null, (s, ev) =>
				{
					if (clickedNode.Text == "未分類")
					{
						dialog.show("「未分類」グループの名前は変更できません。", "エラー");
						return;
					}

					if (dialog.showInput(out string newName, "新しいグループ名を入力してください", clickedNode.Text))
					{
						if (!string.IsNullOrWhiteSpace(newName))
						{
							clickedNode.Text = newName;
							SaveSettings();
						}
					}
				});

				// セパレーター
				menu.Items.Add(new ToolStripSeparator());
			}

			// ✅ 修正: 常に「新しいグループを作成」を表示
			menu.Items.Add("新しいグループを作成", null, (s, ev) =>
			{
				if (!dialog.showInput(out string groupName, "グループ名を入力してください", "新しいグループ")) return;
				if (string.IsNullOrWhiteSpace(groupName)) return;

				var newGroup = new TreeNode(groupName) { Tag = null };
				newGroup.NodeFont = new Font(treeView.Font, FontStyle.Bold);
				newGroup.ForeColor = Color.DarkBlue;
				treeView.Nodes.Add(newGroup);
				SaveSettings();
			});

			menu.Show(treeView, e.Location);
		}
	}

	protected override void Dispose(bool disposing)
	{
		if (disposing)
		{
			refreshTimer?.Stop();
			refreshTimer?.Dispose();
		}
		base.Dispose(disposing);
	}
}
