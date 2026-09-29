using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ShapePath = System.Windows.Shapes.Path;

namespace OfflineDocxCompare
{
    public sealed class MainWindow : Window
    {
        private static readonly Brush PageBackground = BrushFromHex("#F5F8F8");
        private static readonly Brush CardBackground = Brushes.White;
        private static readonly Brush Primary = BrushFromHex("#0D9488");
        private static readonly Brush PrimaryDark = BrushFromHex("#0F5F5A");
        private static readonly Brush Accent = BrushFromHex("#C2410C");
        private static readonly Brush TextPrimary = BrushFromHex("#173B3A");
        private static readonly Brush TextSecondary = BrushFromHex("#55706F");
        private static readonly Brush CardBorderBrush = BrushFromHex("#CDE4E2");
        private static readonly Brush SoftTeal = BrushFromHex("#E9F8F6");

        private readonly TextBlock oldFileName;
        private readonly TextBlock oldFilePath;
        private readonly TextBlock newFileName;
        private readonly TextBlock newFilePath;
        private readonly TextBox outputPathBox;
        private readonly ComboBox colorBox;
        private readonly Button compareButton;
        private readonly ProgressBar progressBar;
        private readonly TextBlock statusTitle;
        private readonly TextBlock statusDetail;
        private readonly Border statusPanel;
        private readonly Button openResultButton;

        private string oldPath;
        private string newPath;
        private string lastOutputPath;
        private bool isBusy;

        public MainWindow()
        {
            Title = "离线文档对比";
            Width = 1040;
            Height = 760;
            MinWidth = 820;
            MinHeight = 650;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            Background = PageBackground;
            FontFamily = new FontFamily("Microsoft YaHei UI, Segoe UI");
            Foreground = TextPrimary;
            UseLayoutRounding = true;
            SnapsToDevicePixels = true;

            Grid root = new Grid();
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            Content = root;

            Border header = BuildHeader();
            Grid.SetRow(header, 0);
            root.Children.Add(header);

            ScrollViewer scrollViewer = new ScrollViewer();
            scrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            scrollViewer.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled;
            scrollViewer.Padding = new Thickness(32, 28, 32, 28);
            Grid.SetRow(scrollViewer, 1);
            root.Children.Add(scrollViewer);

            StackPanel content = new StackPanel();
            content.MaxWidth = 960;
            scrollViewer.Content = content;

            TextBlock intro = new TextBlock();
            intro.Text = "选择旧版本和新版本，程序只比较文字内容，并在新版本副本中标出新增、替换和删除。";
            intro.FontSize = 15;
            intro.Foreground = TextSecondary;
            intro.Margin = new Thickness(0, 0, 0, 22);
            content.Children.Add(intro);

            Grid fileGrid = new Grid();
            fileGrid.ColumnDefinitions.Add(new ColumnDefinition());
            fileGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(20) });
            fileGrid.ColumnDefinitions.Add(new ColumnDefinition());
            content.Children.Add(fileGrid);

            Border oldCard = BuildFileCard(
                "1",
                "旧版本",
                "作为对比基准",
                out oldFileName,
                out oldFilePath,
                delegate { PickFile(true); },
                delegate(string path) { SetFile(true, path); });
            Grid.SetColumn(oldCard, 0);
            fileGrid.Children.Add(oldCard);

            Border newCard = BuildFileCard(
                "2",
                "新版本",
                "输出以此文档为底稿",
                out newFileName,
                out newFilePath,
                delegate { PickFile(false); },
                delegate(string path) { SetFile(false, path); });
            Grid.SetColumn(newCard, 2);
            fileGrid.Children.Add(newCard);

            Border optionsCard = new Border();
            optionsCard.Background = CardBackground;
            optionsCard.BorderBrush = CardBorderBrush;
            optionsCard.BorderThickness = new Thickness(1);
            optionsCard.CornerRadius = new CornerRadius(12);
            optionsCard.Padding = new Thickness(24);
            optionsCard.Margin = new Thickness(0, 20, 0, 0);
            content.Children.Add(optionsCard);

            Grid optionsGrid = new Grid();
            optionsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(210) });
            optionsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(22) });
            optionsGrid.ColumnDefinitions.Add(new ColumnDefinition());
            optionsCard.Child = optionsGrid;

            StackPanel colorPanel = new StackPanel();
            Grid.SetColumn(colorPanel, 0);
            optionsGrid.Children.Add(colorPanel);
            colorPanel.Children.Add(MakeLabel("标注颜色"));
            colorBox = new ComboBox();
            colorBox.Height = 42;
            colorBox.Margin = new Thickness(0, 8, 0, 0);
            colorBox.Padding = new Thickness(8, 3, 8, 3);
            colorBox.ToolTip = "选择写入 Word 的标准高亮颜色";
            AddColorOption("黄色", "yellow", "#FFF200");
            AddColorOption("蓝色", "blue", "#2F75B5");
            AddColorOption("绿色", "green", "#00B050");
            AddColorOption("青色", "cyan", "#00FFFF");
            AddColorOption("粉色", "magenta", "#FF00FF");
            AddColorOption("红色", "red", "#FF0000");
            AddColorOption("浅灰", "lightGray", "#D9D9D9");
            colorBox.SelectedIndex = 0;
            colorPanel.Children.Add(colorBox);

            StackPanel outputPanel = new StackPanel();
            Grid.SetColumn(outputPanel, 2);
            optionsGrid.Children.Add(outputPanel);
            outputPanel.Children.Add(MakeLabel("输出文档"));

            Grid outputGrid = new Grid();
            outputGrid.Margin = new Thickness(0, 8, 0, 0);
            outputGrid.ColumnDefinitions.Add(new ColumnDefinition());
            outputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(10) });
            outputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(104) });
            outputPanel.Children.Add(outputGrid);

            outputPathBox = new TextBox();
            outputPathBox.Height = 42;
            outputPathBox.Padding = new Thickness(11, 8, 11, 8);
            outputPathBox.VerticalContentAlignment = VerticalAlignment.Center;
            outputPathBox.ToolTip = "输出不会覆盖两个源文件";
            outputPathBox.TextChanged += delegate { UpdateCompareButton(); };
            Grid.SetColumn(outputPathBox, 0);
            outputGrid.Children.Add(outputPathBox);

            Button outputBrowse = MakeButton("选择位置", false);
            outputBrowse.Click += delegate { PickOutput(); };
            Grid.SetColumn(outputBrowse, 2);
            outputGrid.Children.Add(outputBrowse);

            compareButton = MakeButton("开始对比并生成文档", true);
            compareButton.Height = 50;
            compareButton.Margin = new Thickness(0, 22, 0, 0);
            compareButton.IsDefault = true;
            compareButton.IsEnabled = false;
            compareButton.IsEnabledChanged += delegate
            {
                compareButton.Foreground = compareButton.IsEnabled ? Brushes.White : BrushFromHex("#657A79");
            };
            compareButton.Foreground = BrushFromHex("#657A79");
            compareButton.Click += async delegate { await StartCompareAsync(); };
            content.Children.Add(compareButton);

            statusPanel = new Border();
            statusPanel.Background = SoftTeal;
            statusPanel.BorderBrush = BrushFromHex("#9ED9D4");
            statusPanel.BorderThickness = new Thickness(1);
            statusPanel.CornerRadius = new CornerRadius(10);
            statusPanel.Padding = new Thickness(18);
            statusPanel.Margin = new Thickness(0, 16, 0, 0);
            statusPanel.Visibility = Visibility.Collapsed;
            content.Children.Add(statusPanel);

            Grid statusGrid = new Grid();
            statusGrid.ColumnDefinitions.Add(new ColumnDefinition());
            statusGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            statusPanel.Child = statusGrid;

            StackPanel statusText = new StackPanel();
            Grid.SetColumn(statusText, 0);
            statusGrid.Children.Add(statusText);
            statusTitle = new TextBlock();
            statusTitle.FontWeight = FontWeights.SemiBold;
            statusTitle.FontSize = 15;
            statusText.Children.Add(statusTitle);
            statusDetail = new TextBlock();
            statusDetail.Margin = new Thickness(0, 5, 0, 0);
            statusDetail.Foreground = TextSecondary;
            statusDetail.TextWrapping = TextWrapping.Wrap;
            statusText.Children.Add(statusDetail);
            progressBar = new ProgressBar();
            progressBar.Height = 6;
            progressBar.Margin = new Thickness(0, 12, 14, 0);
            progressBar.Minimum = 0;
            progressBar.Maximum = 100;
            statusText.Children.Add(progressBar);

            openResultButton = MakeButton("在文件夹中显示", false);
            openResultButton.Margin = new Thickness(18, 0, 0, 0);
            openResultButton.Visibility = Visibility.Collapsed;
            openResultButton.Click += delegate { ShowOutputInFolder(); };
            Grid.SetColumn(openResultButton, 1);
            statusGrid.Children.Add(openResultButton);

            Border footer = BuildFooter();
            Grid.SetRow(footer, 2);
            root.Children.Add(footer);

            RegisterShortcuts();
        }

        private Border BuildHeader()
        {
            Border header = new Border();
            header.Background = Brushes.White;
            header.BorderBrush = CardBorderBrush;
            header.BorderThickness = new Thickness(0, 0, 0, 1);
            header.Padding = new Thickness(32, 20, 32, 20);

            Grid grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition());
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            header.Child = grid;

            Border iconBox = new Border();
            iconBox.Width = 44;
            iconBox.Height = 44;
            iconBox.Background = SoftTeal;
            iconBox.CornerRadius = new CornerRadius(10);
            iconBox.Margin = new Thickness(0, 0, 14, 0);
            ShapePath documentIcon = new ShapePath();
            documentIcon.Data = Geometry.Parse("M10,5 L24,5 L31,12 L31,35 L10,35 Z M24,5 L24,13 L31,13 M15,20 L26,20 M15,25 L26,25 M15,30 L23,30");
            documentIcon.Stroke = Primary;
            documentIcon.StrokeThickness = 2;
            documentIcon.StrokeStartLineCap = PenLineCap.Round;
            documentIcon.StrokeEndLineCap = PenLineCap.Round;
            documentIcon.Stretch = Stretch.None;
            iconBox.Child = documentIcon;
            Grid.SetColumn(iconBox, 0);
            grid.Children.Add(iconBox);

            StackPanel titlePanel = new StackPanel();
            TextBlock title = new TextBlock();
            title.Text = "离线文档对比";
            title.FontSize = 23;
            title.FontWeight = FontWeights.SemiBold;
            titlePanel.Children.Add(title);
            TextBlock subtitle = new TextBlock();
            subtitle.Text = "只比较文字，不比较排版与颜色";
            subtitle.FontSize = 13;
            subtitle.Foreground = TextSecondary;
            subtitle.Margin = new Thickness(0, 4, 0, 0);
            titlePanel.Children.Add(subtitle);
            Grid.SetColumn(titlePanel, 1);
            grid.Children.Add(titlePanel);

            Border offlineBadge = new Border();
            offlineBadge.Background = SoftTeal;
            offlineBadge.BorderBrush = BrushFromHex("#89D2CC");
            offlineBadge.BorderThickness = new Thickness(1);
            offlineBadge.CornerRadius = new CornerRadius(16);
            offlineBadge.Padding = new Thickness(13, 7, 13, 7);
            offlineBadge.VerticalAlignment = VerticalAlignment.Center;
            TextBlock offlineText = new TextBlock();
            offlineText.Text = "完全离线 · 文件不上传";
            offlineText.Foreground = PrimaryDark;
            offlineText.FontWeight = FontWeights.SemiBold;
            offlineText.FontSize = 12;
            offlineBadge.Child = offlineText;
            Grid.SetColumn(offlineBadge, 2);
            grid.Children.Add(offlineBadge);
            return header;
        }

        private Border BuildFileCard(
            string number,
            string title,
            string helper,
            out TextBlock nameBlock,
            out TextBlock pathBlock,
            Action browseAction,
            Action<string> dropAction)
        {
            Border card = new Border();
            card.Background = CardBackground;
            card.BorderBrush = CardBorderBrush;
            card.BorderThickness = new Thickness(1);
            card.CornerRadius = new CornerRadius(12);
            card.Padding = new Thickness(22);
            card.MinHeight = 205;
            card.AllowDrop = true;
            card.ToolTip = "可将一个 .docx 文件拖放到此处";

            StackPanel panel = new StackPanel();
            card.Child = panel;

            Grid heading = new Grid();
            heading.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            heading.ColumnDefinitions.Add(new ColumnDefinition());
            panel.Children.Add(heading);

            Border numberBadge = new Border();
            numberBadge.Width = 30;
            numberBadge.Height = 30;
            numberBadge.Background = Primary;
            numberBadge.CornerRadius = new CornerRadius(15);
            numberBadge.Margin = new Thickness(0, 0, 11, 0);
            TextBlock numberText = new TextBlock();
            numberText.Text = number;
            numberText.Foreground = Brushes.White;
            numberText.FontWeight = FontWeights.Bold;
            numberText.HorizontalAlignment = HorizontalAlignment.Center;
            numberText.VerticalAlignment = VerticalAlignment.Center;
            numberBadge.Child = numberText;
            heading.Children.Add(numberBadge);

            StackPanel headingText = new StackPanel();
            TextBlock titleBlock = new TextBlock();
            titleBlock.Text = title;
            titleBlock.FontSize = 17;
            titleBlock.FontWeight = FontWeights.SemiBold;
            headingText.Children.Add(titleBlock);
            TextBlock helperBlock = new TextBlock();
            helperBlock.Text = helper;
            helperBlock.Foreground = TextSecondary;
            helperBlock.FontSize = 12;
            helperBlock.Margin = new Thickness(0, 3, 0, 0);
            headingText.Children.Add(helperBlock);
            Grid.SetColumn(headingText, 1);
            heading.Children.Add(headingText);

            nameBlock = new TextBlock();
            nameBlock.Text = "尚未选择文件";
            nameBlock.FontWeight = FontWeights.SemiBold;
            nameBlock.FontSize = 14;
            nameBlock.Margin = new Thickness(0, 24, 0, 0);
            nameBlock.TextTrimming = TextTrimming.CharacterEllipsis;
            panel.Children.Add(nameBlock);

            pathBlock = new TextBlock();
            pathBlock.Text = "拖放 DOCX 到这里，或点击下方按钮";
            pathBlock.Foreground = TextSecondary;
            pathBlock.FontSize = 12;
            pathBlock.Margin = new Thickness(0, 6, 0, 0);
            pathBlock.TextTrimming = TextTrimming.CharacterEllipsis;
            panel.Children.Add(pathBlock);

            Button browse = MakeButton("选择 DOCX 文件", false);
            browse.HorizontalAlignment = HorizontalAlignment.Left;
            browse.Margin = new Thickness(0, 20, 0, 0);
            browse.Click += delegate { browseAction(); };
            panel.Children.Add(browse);

            card.DragOver += delegate(object sender, DragEventArgs args)
            {
                string path = GetDroppedDocx(args.Data);
                args.Effects = path == null ? DragDropEffects.None : DragDropEffects.Copy;
                args.Handled = true;
            };
            card.DragEnter += delegate
            {
                card.BorderBrush = Primary;
                card.BorderThickness = new Thickness(2);
            };
            card.DragLeave += delegate
            {
                card.BorderBrush = CardBorderBrush;
                card.BorderThickness = new Thickness(1);
            };
            card.Drop += delegate(object sender, DragEventArgs args)
            {
                card.BorderBrush = CardBorderBrush;
                card.BorderThickness = new Thickness(1);
                string path = GetDroppedDocx(args.Data);
                if (path != null)
                {
                    dropAction(path);
                }
            };
            return card;
        }

        private Border BuildFooter()
        {
            Border footer = new Border();
            footer.Background = Brushes.White;
            footer.BorderBrush = CardBorderBrush;
            footer.BorderThickness = new Thickness(0, 1, 0, 0);
            footer.Padding = new Thickness(32, 12, 32, 12);
            TextBlock text = new TextBlock();
            text.Text = "新增与替换使用所选高亮色 · 旧版删除内容使用红色删除线 · 不对比字体、颜色、背景和版式";
            text.FontSize = 12;
            text.Foreground = TextSecondary;
            text.HorizontalAlignment = HorizontalAlignment.Center;
            text.TextWrapping = TextWrapping.Wrap;
            footer.Child = text;
            return footer;
        }

        private void AddColorOption(string label, string value, string color)
        {
            ComboBoxItem item = new ComboBoxItem();
            item.Tag = value;
            item.Padding = new Thickness(4);
            StackPanel panel = new StackPanel();
            panel.Orientation = Orientation.Horizontal;
            Border swatch = new Border();
            swatch.Width = 18;
            swatch.Height = 18;
            swatch.CornerRadius = new CornerRadius(3);
            swatch.Background = BrushFromHex(color);
            swatch.BorderBrush = BrushFromHex("#8EA5A3");
            swatch.BorderThickness = new Thickness(1);
            swatch.Margin = new Thickness(0, 0, 9, 0);
            panel.Children.Add(swatch);
            TextBlock text = new TextBlock();
            text.Text = label;
            text.VerticalAlignment = VerticalAlignment.Center;
            panel.Children.Add(text);
            item.Content = panel;
            colorBox.Items.Add(item);
        }

        private static TextBlock MakeLabel(string text)
        {
            TextBlock label = new TextBlock();
            label.Text = text;
            label.FontWeight = FontWeights.SemiBold;
            label.FontSize = 14;
            return label;
        }

        private static Button MakeButton(string text, bool primary)
        {
            Button button = new Button();
            button.Content = text;
            button.MinWidth = 104;
            button.Height = 40;
            button.Padding = new Thickness(16, 7, 16, 7);
            button.FontWeight = FontWeights.SemiBold;
            button.Cursor = Cursors.Hand;
            button.Background = primary ? Accent : SoftTeal;
            button.Foreground = primary ? Brushes.White : PrimaryDark;
            button.BorderBrush = primary ? Accent : BrushFromHex("#86CFC9");
            button.BorderThickness = new Thickness(1);
            return button;
        }

        private void PickFile(bool isOld)
        {
            if (isBusy)
            {
                return;
            }
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = isOld ? "选择旧版本 DOCX" : "选择新版本 DOCX";
            dialog.Filter = "Word 文档 (*.docx)|*.docx";
            dialog.Multiselect = false;
            if (dialog.ShowDialog(this) == true)
            {
                SetFile(isOld, dialog.FileName);
            }
        }

        private void SetFile(bool isOld, string path)
        {
            if (!File.Exists(path) || !string.Equals(Path.GetExtension(path), ".docx", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("请选择有效的 .docx 文件。", "文件格式不支持", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (isOld)
            {
                oldPath = Path.GetFullPath(path);
                oldFileName.Text = Path.GetFileName(path);
                oldFilePath.Text = oldPath;
                oldFilePath.ToolTip = oldPath;
            }
            else
            {
                newPath = Path.GetFullPath(path);
                newFileName.Text = Path.GetFileName(path);
                newFilePath.Text = newPath;
                newFilePath.ToolTip = newPath;
                SetDefaultOutputPath();
            }
            UpdateCompareButton();
        }

        private void SetDefaultOutputPath()
        {
            if (string.IsNullOrWhiteSpace(newPath))
            {
                return;
            }
            string directory = Path.GetDirectoryName(newPath);
            string fileName = Path.GetFileNameWithoutExtension(newPath) + "_差异标注.docx";
            outputPathBox.Text = Path.Combine(directory, fileName);
        }

        private void PickOutput()
        {
            if (isBusy)
            {
                return;
            }
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Title = "保存差异标注文档";
            dialog.Filter = "Word 文档 (*.docx)|*.docx";
            dialog.AddExtension = true;
            dialog.DefaultExt = ".docx";
            dialog.OverwritePrompt = true;
            if (!string.IsNullOrWhiteSpace(outputPathBox.Text))
            {
                dialog.InitialDirectory = Path.GetDirectoryName(outputPathBox.Text);
                dialog.FileName = Path.GetFileName(outputPathBox.Text);
            }
            if (dialog.ShowDialog(this) == true)
            {
                outputPathBox.Text = dialog.FileName;
            }
        }

        private async Task StartCompareAsync()
        {
            if (isBusy)
            {
                return;
            }
            if (string.IsNullOrWhiteSpace(oldPath) || string.IsNullOrWhiteSpace(newPath))
            {
                MessageBox.Show("请先选择旧版本和新版本。", "还缺少文件", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string outputPath = outputPathBox.Text.Trim();
            if (File.Exists(outputPath))
            {
                MessageBoxResult overwrite = MessageBox.Show(
                    "输出文件已经存在，是否覆盖？\n\n" + outputPath,
                    "确认覆盖",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);
                if (overwrite != MessageBoxResult.Yes)
                {
                    return;
                }
            }

            ComboBoxItem selectedColor = colorBox.SelectedItem as ComboBoxItem;
            string highlightColor = selectedColor == null ? "yellow" : Convert.ToString(selectedColor.Tag);

            SetBusy(true);
            statusPanel.Visibility = Visibility.Visible;
            statusTitle.Text = "正在比较文档";
            statusDetail.Text = "正在读取文字并定位变化，请稍候…";
            progressBar.Value = 0;
            openResultButton.Visibility = Visibility.Collapsed;

            Progress<int> progress = new Progress<int>(delegate(int value)
            {
                progressBar.Value = value;
                statusDetail.Text = "正在处理，本地进度 " + value + "%";
            });

            try
            {
                CompareResult result = await Task.Run(delegate
                {
                    return DocumentComparer.Compare(oldPath, newPath, outputPath, highlightColor, progress);
                });

                lastOutputPath = result.OutputPath;
                progressBar.Value = 100;
                statusTitle.Text = result.HasDifferences ? "对比完成，已生成标注文档" : "对比完成，未发现文字差异";
                statusDetail.Text = BuildResultSummary(result);
                openResultButton.Visibility = Visibility.Visible;
            }
            catch (Exception exception)
            {
                statusTitle.Text = "对比未完成";
                statusDetail.Text = exception.Message;
                progressBar.Value = 0;
                MessageBox.Show(
                    exception.Message,
                    "无法完成对比",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                SetBusy(false);
            }
        }

        private static string BuildResultSummary(CompareResult result)
        {
            string summary = string.Format(
                "变化段落 {0} 处（新增 {1}、替换 {2}），高亮新文字 {3} 个字符；旧版删除文字 {4} 个字符已用红色删除线保留。",
                result.ChangedParagraphs,
                result.InsertedParagraphs,
                result.ReplacedParagraphs,
                result.HighlightedCharacters,
                result.DeletedCharactersShown);
            if (result.DeletedParagraphs > 0)
            {
                summary += string.Format(
                    " 已将旧版纯删除段落 {0} 处以红色删除线插入副本。",
                    result.DeletedParagraphs);
            }
            return summary;
        }

        private void SetBusy(bool busy)
        {
            isBusy = busy;
            oldFileName.IsEnabled = !busy;
            newFileName.IsEnabled = !busy;
            outputPathBox.IsEnabled = !busy;
            colorBox.IsEnabled = !busy;
            compareButton.Content = busy ? "正在生成，请稍候…" : "开始对比并生成文档";
            UpdateCompareButton();
            Cursor = busy ? Cursors.Wait : Cursors.Arrow;
        }

        private void UpdateCompareButton()
        {
            if (compareButton == null)
            {
                return;
            }
            bool ready = !string.IsNullOrWhiteSpace(oldPath) &&
                         !string.IsNullOrWhiteSpace(newPath) &&
                         !string.IsNullOrWhiteSpace(outputPathBox.Text);
            compareButton.IsEnabled = !isBusy && ready;
            compareButton.Foreground = compareButton.IsEnabled ? Brushes.White : BrushFromHex("#657A79");
            if (!isBusy)
            {
                compareButton.Content = ready ? "开始对比并生成文档" : "请先选择旧版本和新版本";
            }
        }

        private void ShowOutputInFolder()
        {
            if (string.IsNullOrWhiteSpace(lastOutputPath) || !File.Exists(lastOutputPath))
            {
                return;
            }
            Process.Start("explorer.exe", "/select,\"" + lastOutputPath + "\"");
        }

        private void RegisterShortcuts()
        {
            RoutedCommand oldCommand = new RoutedCommand();
            oldCommand.InputGestures.Add(new KeyGesture(Key.O, ModifierKeys.Control));
            CommandBindings.Add(new CommandBinding(oldCommand, delegate { PickFile(true); }));
            InputBindings.Add(new InputBinding(oldCommand, new KeyGesture(Key.O, ModifierKeys.Control)));

            RoutedCommand newCommand = new RoutedCommand();
            newCommand.InputGestures.Add(new KeyGesture(Key.O, ModifierKeys.Control | ModifierKeys.Shift));
            CommandBindings.Add(new CommandBinding(newCommand, delegate { PickFile(false); }));
            InputBindings.Add(new InputBinding(newCommand, new KeyGesture(Key.O, ModifierKeys.Control | ModifierKeys.Shift)));

            RoutedCommand compareCommand = new RoutedCommand();
            compareCommand.InputGestures.Add(new KeyGesture(Key.R, ModifierKeys.Control));
            CommandBindings.Add(new CommandBinding(compareCommand, async delegate { await StartCompareAsync(); }));
            InputBindings.Add(new InputBinding(compareCommand, new KeyGesture(Key.R, ModifierKeys.Control)));
        }

        private static string GetDroppedDocx(IDataObject data)
        {
            if (data == null || !data.GetDataPresent(DataFormats.FileDrop))
            {
                return null;
            }
            string[] paths = data.GetData(DataFormats.FileDrop) as string[];
            if (paths == null || paths.Length == 0)
            {
                return null;
            }
            foreach (string path in paths)
            {
                if (File.Exists(path) && string.Equals(Path.GetExtension(path), ".docx", StringComparison.OrdinalIgnoreCase))
                {
                    return path;
                }
            }
            return null;
        }

        private static SolidColorBrush BrushFromHex(string value)
        {
            return new SolidColorBrush((Color)ColorConverter.ConvertFromString(value));
        }
    }
}
