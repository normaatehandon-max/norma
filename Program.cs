using ClosedXML.Excel;
using System.Globalization;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace CedulaIngresos;

internal static class Program
{
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}

public sealed class MainForm : Form
{
    private string? _pdfPath;
    private string? _auxPath;
    private List<Deposit> _deposits = new();

    private readonly Label _lblPdf = new();
    private readonly Label _lblAux = new();
    private readonly Label _status = new();
    private readonly RichTextBox _summary = new();
    private readonly ProgressBar _progress = new();

    public MainForm()
    {
        Text = "CÉDULA DE INGRESOS";
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(820, 650);
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        BackColor = Color.FromArgb(37, 37, 37);
        Font = new Font("Segoe UI", 10);

        var title = new Label
        {
            Text = "CÉDULA DE INGRESOS",
            AutoSize = true,
            Font = new Font("Segoe UI", 20, FontStyle.Bold),
            ForeColor = Color.FromArgb(80, 227, 194),
            BackColor = BackColor,
            Location = new Point(250, 22)
        };
        Controls.Add(title);

        var subtitle = new Label
        {
            Text = "Estado de cuenta + Auxiliar 2150 → Excel",
            AutoSize = true,
            ForeColor = Color.LightGray,
            BackColor = BackColor,
            Location = new Point(275, 62)
        };
        Controls.Add(subtitle);

        int left = 70, width = 680, y = 105;

        var btnPdf = MakeButton("1. CARGAR ESTADO DE CUENTA (PDF)", left, y, width, PickPdf);
        Controls.Add(btnPdf);
        y += 48;

        _lblPdf.Text = "Ningún PDF cargado";
        _lblPdf.SetBounds(left, y, width, 22);
        _lblPdf.ForeColor = Color.Silver;
        _lblPdf.BackColor = BackColor;
        Controls.Add(_lblPdf);
        y += 38;

        var btnAux = MakeButton("2. CARGAR AUXILIAR 2150 (EXCEL)", left, y, width, PickAux);
        Controls.Add(btnAux);
        y += 48;

        _lblAux.Text = "Ningún auxiliar cargado";
        _lblAux.SetBounds(left, y, width, 22);
        _lblAux.ForeColor = Color.Silver;
        _lblAux.BackColor = BackColor;
        Controls.Add(_lblAux);
        y += 42;

        var btnProcess = MakeButton("3. PROCESAR Y CONCILIAR", left, y, width, async (_, __) => await ProcessAsync());
        Controls.Add(btnProcess);
        y += 58;

        var btnGenerate = MakeButton("4. GENERAR CÉDULA DE INGRESOS", left, y, width, async (_, __) => await GenerateAsync());
        btnGenerate.BackColor = Color.FromArgb(25, 135, 84);
        Controls.Add(btnGenerate);
        y += 64;

        _status.Text = "Listo para procesar";
        _status.SetBounds(left, y, width, 24);
        _status.ForeColor = Color.White;
        _status.BackColor = BackColor;
        Controls.Add(_status);
        y += 30;

        _progress.SetBounds(left, y, width, 18);
        _progress.Style = ProgressBarStyle.Marquee;
        _progress.MarqueeAnimationSpeed = 0;
        Controls.Add(_progress);
        y += 38;

        var lblSummary = new Label
        {
            Text = "Resumen:",
            AutoSize = true,
            ForeColor = Color.White,
            BackColor = BackColor,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Location = new Point(left, y)
        };
        Controls.Add(lblSummary);
        y += 28;

        _summary.SetBounds(left, y, width, 170);
        _summary.BackColor = Color.FromArgb(51, 51, 51);
        _summary.ForeColor = Color.White;
        _summary.ReadOnly = true;
        _summary.BorderStyle = BorderStyle.None;
        _summary.Text = "Aún no hay archivos procesados.";
        Controls.Add(_summary);
    }

    private Button MakeButton(string text, int x, int y, int w, EventHandler click)
    {
        var b = new Button
        {
            Text = text,
            BackColor = Color.FromArgb(52, 88, 170),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 11, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        b.SetBounds(x, y, w, 42);
        b.FlatAppearance.BorderSize = 0;
        b.Click += click;
        return b;
    }

    private void PickPdf(object? sender, EventArgs e)
    {
        using var dlg = new OpenFileDialog
        {
            Title = "Seleccionar estado de cuenta",
            Filter = "PDF (*.pdf)|*.pdf"
        };
        if (dlg.ShowDialog() == DialogResult.OK)
        {
            _pdfPath = dlg.FileName;
            _lblPdf.Text = Path.GetFileName(_pdfPath);
        }
    }

    private void PickAux(object? sender, EventArgs e)
    {
        using var dlg = new OpenFileDialog
        {
            Title = "Seleccionar Auxiliar 2150",
            Filter = "Excel (*.xlsx;*.xls)|*.xlsx;*.xls"
        };
        if (dlg.ShowDialog() == DialogResult.OK)
        {
            _auxPath = dlg.FileName;
            _lblAux.Text = Path.GetFileName(_auxPath);
        }
    }

    private async Task ProcessAsync()
    {
        if (string.IsNullOrWhiteSpace(_pdfPath) || string.IsNullOrWhiteSpace(_auxPath))
        {
            MessageBox.Show("Carga primero el estado de cuenta PDF y el Auxiliar 2150.",
                "Cédula de Ingresos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            Busy(true, "Procesando...");
            var result = await Task.Run(() =>
            {
                var text = PdfText(_pdfPath!);
                var deposits = ParseBbva(text);
                if (deposits.Count == 0) deposits = ParseGeneric(text);
                if (deposits.Count == 0)
                    throw new InvalidOperationException("No pude detectar depósitos en este estado de cuenta.");

                var aux = LoadAux(_auxPath!);
                Match(deposits, aux);
                return deposits;
            });

            _deposits = result;
            var total = _deposits.Sum(x => x.Amount);
            var matched = _deposits.Count(x => !string.IsNullOrWhiteSpace(x.Uuid));
            var review = _deposits.Count(x => x.Status.Contains("REVISAR", StringComparison.OrdinalIgnoreCase));

            _summary.Text =
                $"Depósitos detectados: {_deposits.Count}\r\n" +
                $"Total depósitos: ${total:N2}\r\n" +
                $"Con UUID automático: {matched}\r\n" +
                $"Por revisar: {review}\r\n\r\n" +
                "Cada depósito se conserva individualmente, sin agrupar ni modificar.";
            _status.Text = "Listo para generar la cédula.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Cédula de Ingresos", MessageBoxButtons.OK, MessageBoxIcon.Error);
            _status.Text = "Error al procesar.";
        }
        finally
        {
            Busy(false, _status.Text);
        }
    }

    private async Task GenerateAsync()
    {
        if (_deposits.Count == 0)
        {
            MessageBox.Show("Primero pulsa PROCESAR Y CONCILIAR.",
                "Cédula de Ingresos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        using var dlg = new SaveFileDialog
        {
            Title = "Guardar Cédula de Ingresos",
            Filter = "Excel (*.xlsx)|*.xlsx",
            FileName = "CEDULA_INGRESOS.xlsx",
            AddExtension = true,
            DefaultExt = "xlsx"
        };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        try
        {
            Busy(true, "Generando Excel...");
            await Task.Run(() => Export(dlg.FileName, _deposits));
            MessageBox.Show($"Cédula generada correctamente:\r\n{dlg.FileName}",
                "Cédula de Ingresos", MessageBoxButtons.OK, MessageBoxIcon.Information);
            _status.Text = "Cédula generada correctamente.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Cédula de Ingresos", MessageBoxButtons.OK, MessageBoxIcon.Error);
            _status.Text = "Error al generar Excel.";
        }
        finally
        {
            Busy(false, _status.Text);
        }
    }

    private void Busy(bool yes, string status)
    {
        _status.Text = status;
        _progress.MarqueeAnimationSpeed = yes ? 30 : 0;
        UseWaitCursor = yes;
        Application.DoEvents();
    }

    private static string PdfText(string path)
    {
        using var doc = PdfDocument.Open(path);
        return string.Join("\n", doc.GetPages().Select(p => p.Text));
    }

    private static List<Deposit> ParseBbva(string text)
    {
        var list = new List<Deposit>();
        var lines = text.Replace("\r", "\n")
            .Split('\n')
            .Select(x => Regex.Replace(x, @"\s+", " ").Trim())
            .Where(x => x.Length > 0)
            .ToArray();

        int year = 2025;
        var ym = Regex.Match(text, @"DEL\s+\d{2}/\d{2}/(?<y>\d{4})\s+AL", RegexOptions.IgnoreCase);
        if (ym.Success) year = int.Parse(ym.Groups["y"].Value);

        var monMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["ENE"]=1,["FEB"]=2,["MAR"]=3,["ABR"]=4,["MAY"]=5,["JUN"]=6,
            ["JUL"]=7,["AGO"]=8,["SEP"]=9,["OCT"]=10,["NOV"]=11,["DIC"]=12
        };

        var start = new Regex(@"^(?<day>\d{2})/(?<mon>[A-Z]{3})\s+\d{2}/[A-Z]{3}\s+(?<code>T20|C07|N16|W02|AA7)\s+(?<rest>.+)$");
        var money = new Regex(@"(?<!\d)(\d{1,3}(?:,\d{3})*\.\d{2})(?!\d)");

        for (int i = 0; i < lines.Length; i++)
        {
            var m = start.Match(lines[i]);
            if (!m.Success) continue;

            var mm = money.Match(m.Groups["rest"].Value);
            if (!mm.Success) continue;

            var amountText = mm.Groups[1].Value;
            var amount = decimal.Parse(amountText, NumberStyles.Number, CultureInfo.GetCultureInfo("en-US"));
            var rest = m.Groups["rest"].Value;
            var concept = rest[..mm.Index].Trim();

            var extra = new List<string>();
            for (int j = i + 1; j < Math.Min(i + 5, lines.Length); j++)
            {
                var n = lines[j];
                if (Regex.IsMatch(n, @"^\d{2}/[A-Z]{3}\s+")) break;
                if (n.StartsWith("No. Cuenta") || n.StartsWith("BBVA MEXICO") ||
                    n.StartsWith("PAGINA") || n.StartsWith("MAESTRA PYME") ||
                    n.StartsWith("FECHA SALDO")) break;
                extra.Add(n);
            }
            if (extra.Count > 0) concept = (concept + " " + string.Join(" ", extra)).Trim();

            int mon = monMap.TryGetValue(m.Groups["mon"].Value, out var mv) ? mv : 1;
            list.Add(new Deposit
            {
                Date = new DateTime(year, mon, int.Parse(m.Groups["day"].Value)),
                Amount = amount,
                Concept = concept,
                Bank = "BBVA"
            });
        }
        return list;
    }

    private static List<Deposit> ParseGeneric(string text)
    {
        var list = new List<Deposit>();
        var money = new Regex(@"(?<!\d)(\d{1,3}(?:,\d{3})*\.\d{2})(?!\d)");
        foreach (var raw in text.Split('\n'))
        {
            var line = Regex.Replace(raw, @"\s+", " ").Trim();
            var dm = Regex.Match(line, @"\b(\d{2})[/-](\d{2})[/-](\d{4})\b");
            var amounts = money.Matches(line);
            if (!dm.Success || amounts.Count == 0) continue;

            var a = amounts[^1].Groups[1].Value;
            list.Add(new Deposit
            {
                Date = new DateTime(int.Parse(dm.Groups[3].Value), int.Parse(dm.Groups[2].Value), int.Parse(dm.Groups[1].Value)),
                Amount = decimal.Parse(a, NumberStyles.Number, CultureInfo.GetCultureInfo("en-US")),
                Concept = line,
                Bank = "OTRO"
            });
        }
        return list;
    }

    private static List<AuxRow> LoadAux(string path)
    {
        var rows = new List<AuxRow>();
        using var wb = new XLWorkbook(path);
        var ws = wb.Worksheet(1);
        var used = ws.RangeUsed();
        if (used is null) return rows;

        var uuidRx = new Regex(@"\b[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\b");
        var folioRx = new Regex(@"\bF-\s*([0-9]+)\b", RegexOptions.IgnoreCase);

        foreach (var row in used.Rows())
        {
            var vals = row.Cells().Select(c => c.GetFormattedString().Trim()).ToArray();
            var joined = string.Join(" | ", vals);
            var um = uuidRx.Match(joined);
            if (!um.Success) continue;

            var fm = folioRx.Match(joined);
            var concept = vals.Length >= 5 ? vals[4] : joined;
            var parts = concept.Split('|').Select(x => x.Trim()).ToArray();
            var client = parts.Length >= 2 ? parts[1] : "";

            rows.Add(new AuxRow
            {
                Uuid = um.Value.ToUpperInvariant(),
                Folio = fm.Success ? fm.Groups[1].Value : "",
                Client = client,
                Text = joined
            });
        }
        return rows;
    }

    private static void Match(List<Deposit> deposits, List<AuxRow> aux)
    {
        foreach (var d in deposits)
        {
            var c = d.Concept.ToUpperInvariant();
            d.Classification = "NO IDENTIFICADO";
            d.Status = "REVISAR - SIN UUID ÚNICO";

            if (c.Contains("REEMBOLS"))
            {
                d.Classification = "REEMBOLSABLES";
                d.Status = "CLASIFICADO POR CONCEPTO";
            }
            else if (c.Contains("DEVOLU"))
            {
                d.Classification = "DEVOLUCION";
                d.Status = "CLASIFICADO POR CONCEPTO";
            }
            else if (c.Contains("TRASPAS"))
            {
                d.Classification = "TRASPASO";
                d.Status = "CLASIFICADO POR CONCEPTO";
            }

            var candidates = aux
                .Where(a => !string.IsNullOrWhiteSpace(a.Client) &&
                            a.Client.Length >= 4 &&
                            c.Contains(a.Client.ToUpperInvariant()))
                .GroupBy(a => a.Uuid)
                .Select(g => g.First())
                .ToList();

            if (candidates.Count == 1)
            {
                d.Uuid = candidates[0].Uuid;
                d.Folio = candidates[0].Folio;
                if (d.Classification == "NO IDENTIFICADO")
                    d.Classification = "COBRANZA tasa 16";
                d.Status = "AUTO - UUID ÚNICO";
            }
        }
    }

    private static void Export(string outPath, List<Deposit> deposits)
    {
        var template = Path.Combine(AppContext.BaseDirectory, "cedulas INGRESOS(3).xlsx");
        if (!File.Exists(template))
            throw new FileNotFoundException("No encontré la plantilla cedulas INGRESOS(3).xlsx.");

        File.Copy(template, outPath, true);

        using var wb = new XLWorkbook(outPath);
        var ws = wb.Worksheet("INGRESOS");

        ws.Cell("O4").Value = "UUID";
        ws.Cell("R4").Value = "ESTATUS CRUCE";
        ws.Range("B5:R2004").Clear(XLClearOptions.Contents);

        int row = 5;
        foreach (var d in deposits)
        {
            ws.Cell(row, 2).Value = d.Date;
            ws.Cell(row, 3).Value = d.Concept;
            ws.Cell(row, 5).Value = d.Amount;

            int? classCol = d.Classification switch
            {
                "COBRANZA tasa 16" => 6,
                "cobranza 0" => 7,
                "INTERESES" => 8,
                "REEMBOLSABLES" => 9,
                "DEVOLUCION" => 10,
                "NO IDENTIFICADO" => 11,
                "TRASPASO" => 12,
                "TRASPASO INVERSION" => 13,
                "PRESTAMO" => 14,
                _ => null
            };
            if (classCol.HasValue)
                ws.Cell(row, classCol.Value).Value = d.Amount;

            ws.Cell(row, 15).Value = d.Uuid;
            ws.Cell(row, 16).Value = d.Folio;
            ws.Cell(row, 17).FormulaA1 = $"=E{row}-SUM(F{row}:N{row})";
            ws.Cell(row, 18).Value = d.Status;
            row++;
        }

        ws.Column(2).Style.DateFormat.Format = "dd/mm/yyyy";
        foreach (var col in new[] { 4,5,6,7,8,9,10,11,12,13,14,17 })
            ws.Column(col).Style.NumberFormat.Format = "#,##0.00";

        ws.Column(3).Width = 55;
        ws.Column(15).Width = 40;
        ws.Column(16).Width = 16;
        ws.Column(18).Width = 28;

        wb.Save();
    }

    private sealed class Deposit
    {
        public DateTime Date { get; set; }
        public decimal Amount { get; set; }
        public string Concept { get; set; } = "";
        public string Bank { get; set; } = "";
        public string Uuid { get; set; } = "";
        public string Folio { get; set; } = "";
        public string Classification { get; set; } = "";
        public string Status { get; set; } = "";
    }

    private sealed class AuxRow
    {
        public string Uuid { get; set; } = "";
        public string Folio { get; set; } = "";
        public string Client { get; set; } = "";
        public string Text { get; set; } = "";
    }
}
