using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using UglyToad.PdfPig;

namespace CedulaIngresos
{
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
        private string _pdfPath = "";
        private string _auxPath = "";
        private string _compPath = "";

        private readonly Label _lblPdf = new Label();
        private readonly Label _lblAux = new Label();
        private readonly Label _lblComp = new Label();
        private readonly Label _status = new Label();
        private readonly RichTextBox _summary = new RichTextBox();
        private readonly ProgressBar _progress = new ProgressBar();

        private List<BankDeposit> _deposits = new List<BankDeposit>();
        private List<AuxItem> _auxItems = new List<AuxItem>();
        private ComplementBook _complements = new ComplementBook();

        public MainForm()
        {
            Text = "CÉDULA DE INGRESOS";
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(850, 720);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            BackColor = Color.FromArgb(37, 37, 37);
            Font = new Font("Segoe UI", 10);

            Label title = new Label();
            title.Text = "CÉDULA DE INGRESOS";
            title.AutoSize = true;
            title.Font = new Font("Segoe UI", 20, FontStyle.Bold);
            title.ForeColor = Color.FromArgb(80, 227, 194);
            title.BackColor = BackColor;
            title.Location = new Point(265, 20);
            Controls.Add(title);

            Label subtitle = new Label();
            subtitle.Text = "Banco + Auxiliar 2150 + Complementos de pago → Excel";
            subtitle.AutoSize = true;
            subtitle.ForeColor = Color.LightGray;
            subtitle.BackColor = BackColor;
            subtitle.Location = new Point(225, 58);
            Controls.Add(subtitle);

            int left = 70;
            int width = 710;
            int y = 95;

            Controls.Add(MakeButton("1. CARGAR ESTADO DE CUENTA (PDF)", left, y, width, PickPdf));
            y += 46;
            SetupFileLabel(_lblPdf, "Ningún estado de cuenta cargado", left, y, width);
            y += 36;

            Controls.Add(MakeButton("2. CARGAR AUXILIAR 2150 (EXCEL)", left, y, width, PickAux));
            y += 46;
            SetupFileLabel(_lblAux, "Ningún Auxiliar 2150 cargado", left, y, width);
            y += 36;

            Controls.Add(MakeButton("3. CARGAR COMPLEMENTOS DE PAGO (EXCEL)", left, y, width, PickComp));
            y += 46;
            SetupFileLabel(_lblComp, "Ningún archivo de complementos cargado", left, y, width);
            y += 42;

            Controls.Add(MakeButton("4. PROCESAR Y CONCILIAR", left, y, width, ProcessClick));
            y += 56;

            Button btnGenerate = MakeButton("5. GENERAR CÉDULA DE INGRESOS", left, y, width, GenerateClick);
            btnGenerate.BackColor = Color.FromArgb(25, 135, 84);
            Controls.Add(btnGenerate);
            y += 62;

            _status.Text = "Listo para procesar";
            _status.SetBounds(left, y, width, 24);
            _status.ForeColor = Color.White;
            _status.BackColor = BackColor;
            Controls.Add(_status);
            y += 28;

            _progress.SetBounds(left, y, width, 18);
            _progress.Style = ProgressBarStyle.Marquee;
            _progress.MarqueeAnimationSpeed = 0;
            Controls.Add(_progress);
            y += 34;

            Label lblSummary = new Label();
            lblSummary.Text = "Resumen de conciliación:";
            lblSummary.AutoSize = true;
            lblSummary.ForeColor = Color.White;
            lblSummary.BackColor = BackColor;
            lblSummary.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lblSummary.Location = new Point(left, y);
            Controls.Add(lblSummary);
            y += 27;

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
            Button b = new Button();
            b.Text = text;
            b.BackColor = Color.FromArgb(52, 88, 170);
            b.ForeColor = Color.White;
            b.FlatStyle = FlatStyle.Flat;
            b.Font = new Font("Segoe UI", 11, FontStyle.Bold);
            b.Cursor = Cursors.Hand;
            b.SetBounds(x, y, w, 40);
            b.FlatAppearance.BorderSize = 0;
            b.Click += click;
            return b;
        }

        private void SetupFileLabel(Label lbl, string text, int x, int y, int width)
        {
            lbl.Text = text;
            lbl.SetBounds(x, y, width, 22);
            lbl.ForeColor = Color.Silver;
            lbl.BackColor = BackColor;
            Controls.Add(lbl);
        }

        private void PickPdf(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Title = "Estado de cuenta";
                dlg.Filter = "PDF (*.pdf)|*.pdf";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    _pdfPath = dlg.FileName;
                    _lblPdf.Text = Path.GetFileName(_pdfPath);
                }
            }
        }

        private void PickAux(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Title = "Auxiliar 2150";
                dlg.Filter = "Excel (*.xlsx)|*.xlsx";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    _auxPath = dlg.FileName;
                    _lblAux.Text = Path.GetFileName(_auxPath);
                }
            }
        }

        private void PickComp(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Title = "Complementos de pago";
                dlg.Filter = "Excel (*.xlsx)|*.xlsx";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    _compPath = dlg.FileName;
                    _lblComp.Text = Path.GetFileName(_compPath);
                }
            }
        }

        private async void ProcessClick(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(_pdfPath) ||
                String.IsNullOrWhiteSpace(_auxPath) ||
                String.IsNullOrWhiteSpace(_compPath))
            {
                MessageBox.Show(
                    "Carga los 3 archivos:\r\n1) Estado de cuenta\r\n2) Auxiliar 2150\r\n3) Complementos de pago.",
                    "Cédula de Ingresos",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            try
            {
                Busy(true, "Procesando y amarrando información...");

                await Task.Run(delegate
                {
                    string bankText = ReadPdf(_pdfPath);
                    _deposits = ParseBbvaDeposits(bankText);
                    _auxItems = LoadAuxiliar(_auxPath);
                    _complements = LoadComplements(_compPath);

                    LinkComplements(_auxItems, _complements);
                    MatchDeposits(_deposits, _auxItems);
                });

                decimal total = _deposits.Sum(x => x.Amount);
                int matched = _deposits.Count(x => x.Income != null);
                int withComp = _deposits.Count(x => x.Income != null &&
                    x.Income.Items.Any(i => i.Complement != null));

                _summary.Text =
                    "Depósitos/abonos: " + _deposits.Count + "\r\n" +
                    "Total banco: $" + total.ToString("N2") + "\r\n" +
                    "Amarrados a ingreso 2150: " + matched + "\r\n" +
                    "Con complemento localizado: " + withComp + "\r\n" +
                    "Por revisar: " + (_deposits.Count - matched) + "\r\n\r\n" +
                    "El depósito del banco se conserva una sola vez.";

                _status.Text = "Listo para generar la cédula.";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Cédula de Ingresos",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                _status.Text = "Error al procesar.";
            }
            finally
            {
                Busy(false, _status.Text);
            }
        }

        private async void GenerateClick(object sender, EventArgs e)
        {
            if (_deposits.Count == 0)
            {
                MessageBox.Show("Primero pulsa PROCESAR Y CONCILIAR.",
                    "Cédula de Ingresos",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Title = "Guardar Cédula de Ingresos";
                dlg.Filter = "Excel (*.xlsx)|*.xlsx";
                dlg.FileName = "CEDULA_INGRESOS.xlsx";
                dlg.AddExtension = true;
                dlg.DefaultExt = "xlsx";

                if (dlg.ShowDialog() != DialogResult.OK)
                    return;

                try
                {
                    Busy(true, "Generando Excel...");
                    await Task.Run(delegate { ExportExcel(dlg.FileName); });

                    MessageBox.Show("Cédula generada correctamente:\r\n" + dlg.FileName,
                        "Cédula de Ingresos",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                    _status.Text = "Cédula generada correctamente.";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Cédula de Ingresos",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    _status.Text = "Error al generar Excel.";
                }
                finally
                {
                    Busy(false, _status.Text);
                }
            }
        }

        private void Busy(bool yes, string text)
        {
            _status.Text = text;
            _progress.MarqueeAnimationSpeed = yes ? 30 : 0;
            UseWaitCursor = yes;
            Application.DoEvents();
        }

        private static string ReadPdf(string path)
        {
            StringBuilder sb = new StringBuilder();
            using (PdfDocument doc = PdfDocument.Open(path))
            {
                foreach (var page in doc.GetPages())
                    sb.AppendLine(page.Text);
            }
            return sb.ToString();
        }

        private static List<BankDeposit> ParseBbvaDeposits(string text)
        {
            List<BankDeposit> result = new List<BankDeposit>();
            string normalized = text.Replace("\r", "\n");
            normalized = Regex.Replace(normalized, @"[ \t]+", " ");

            int year = DateTime.Now.Year;
            Match yearMatch = Regex.Match(normalized,
                @"DEL\s+\d{2}/\d{2}/(?<y>\d{4})\s+AL",
                RegexOptions.IgnoreCase);

            if (yearMatch.Success)
                year = Int32.Parse(yearMatch.Groups["y"].Value);

            Dictionary<string, int> months =
                new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            months["ENE"] = 1;
            months["FEB"] = 2;
            months["MAR"] = 3;
            months["ABR"] = 4;
            months["MAY"] = 5;
            months["JUN"] = 6;
            months["JUL"] = 7;
            months["AGO"] = 8;
            months["SEP"] = 9;
            months["OCT"] = 10;
            months["NOV"] = 11;
            months["DIC"] = 12;

            Regex startRx = new Regex(
                @"(?m)^(?<day>\d{2})/(?<mon>[A-Z]{3})\s+\d{2}/[A-Z]{3}\s+" +
                @"(?<code>T20|C07|N16|W02|AA7)\s+(?<first>[^\n]+)");

            MatchCollection starts = startRx.Matches(normalized);

            for (int i = 0; i < starts.Count; i++)
            {
                Match m = starts[i];
                int end = (i + 1 < starts.Count) ? starts[i + 1].Index : normalized.Length;
                string block = normalized.Substring(m.Index, end - m.Index);

                string[] stopTexts =
                {
                    "BBVA MEXICO, S.A.",
                    "Total de Movimientos",
                    "No. Cuenta "
                };

                foreach (string stop in stopTexts)
                {
                    int p = block.IndexOf(stop, StringComparison.OrdinalIgnoreCase);
                    if (p > 0)
                        block = block.Substring(0, p);
                }

                List<string> lines = block.Split('\n')
                    .Select(x => Regex.Replace(x, @"\s+", " ").Trim())
                    .Where(x => x.Length > 0)
                    .ToList();

                if (lines.Count == 0)
                    continue;

                Match money = Regex.Match(lines[0],
                    @"(?<!\d)(\d{1,3}(?:,\d{3})*\.\d{2})(?!\d)");

                if (!money.Success)
                    continue;

                decimal amount = Decimal.Parse(
                    money.Groups[1].Value,
                    NumberStyles.Number,
                    CultureInfo.GetCultureInfo("en-US"));

                string first = Regex.Replace(
                    lines[0],
                    @"^\d{2}/[A-Z]{3}\s+\d{2}/[A-Z]{3}\s+(T20|C07|N16|W02|AA7)\s+",
                    "",
                    RegexOptions.IgnoreCase);

                int moneyPos = first.IndexOf(money.Groups[1].Value, StringComparison.Ordinal);
                if (moneyPos >= 0)
                    first = first.Substring(0, moneyPos).Trim();

                List<string> conceptLines = new List<string>();
                if (first.Length > 0)
                    conceptLines.Add(first);

                for (int j = 1; j < lines.Count; j++)
                {
                    string ln = lines[j];

                    if (ln.StartsWith("BBVA MEXICO", StringComparison.OrdinalIgnoreCase) ||
                        ln.StartsWith("PAGINA", StringComparison.OrdinalIgnoreCase) ||
                        ln.StartsWith("MAESTRA PYME", StringComparison.OrdinalIgnoreCase) ||
                        ln.StartsWith("FECHA SALDO", StringComparison.OrdinalIgnoreCase) ||
                        ln.StartsWith("No. Cuenta", StringComparison.OrdinalIgnoreCase))
                        break;

                    if (Regex.IsMatch(ln,
                        @"^\d{1,3}(?:,\d{3})*\.\d{2}(\s+\d{1,3}(?:,\d{3})*\.\d{2})?$"))
                        continue;

                    conceptLines.Add(ln);
                }

                int month = months.ContainsKey(m.Groups["mon"].Value)
                    ? months[m.Groups["mon"].Value] : 1;

                BankDeposit d = new BankDeposit();
                d.Order = result.Count + 1;
                d.Date = new DateTime(year, month, Int32.Parse(m.Groups["day"].Value));
                d.Amount = amount;
                d.Concept = String.Join(Environment.NewLine, conceptLines).Trim();
                result.Add(d);
            }

            if (result.Count == 0)
                throw new InvalidOperationException(
                    "No pude detectar depósitos/abonos BBVA en el estado de cuenta.");

            return result;
        }

        private static List<AuxItem> LoadAuxiliar(string path)
        {
            List<AuxItem> result = new List<AuxItem>();

            using (XLWorkbook wb = new XLWorkbook(path))
            {
                foreach (IXLWorksheet ws in wb.Worksheets)
                {
                    IXLRange used = ws.RangeUsed();
                    if (used == null)
                        continue;

                    int headerRow = FindHeaderRow(ws, 20);
                    Dictionary<string, int> h = BuildHeaderMap(ws, headerRow);

                    int cIngreso = FindColumn(h, new string[] { "INGRESO", "NO INGRESO", "NUM INGRESO" });
                    int cFecha = FindColumn(h, new string[] { "FECHA" });
                    int cUuid = FindColumn(h, new string[] { "UUID" });
                    int cCliente = FindColumn(h, new string[] { "CLIENTE" });
                    int cFactura = FindColumn(h, new string[] { "FACTURA", "FOLIO" });
                    int cIva4 = FindColumn(h, new string[] { "IVA 4", "IVA4" });
                    int cIva16 = FindColumn(h, new string[] { "IVA 16", "IVA16" });
                    int cConcepto = FindColumn(h, new string[] { "CONCEPTO" });
                    int cDebe = FindColumn(h, new string[] { "DEBE" });
                    int cHaber = FindColumn(h, new string[] { "HABER" });

                    foreach (IXLRow row in ws.RowsUsed())
                    {
                        if (row.RowNumber() <= headerRow)
                            continue;

                        string all = String.Join(" | ",
                            row.CellsUsed().Select(c => c.GetFormattedString().Trim()));

                        string uuid = "";
                        if (cUuid > 0)
                            uuid = row.Cell(cUuid).GetFormattedString().Trim();

                        if (String.IsNullOrWhiteSpace(uuid))
                            uuid = ExtractUuid(all);

                        if (String.IsNullOrWhiteSpace(uuid))
                            continue;

                        AuxItem item = new AuxItem();
                        item.InvoiceUuid = uuid.ToUpperInvariant();

                        if (cIngreso > 0)
                            item.IncomeNo = row.Cell(cIngreso).GetFormattedString().Trim();

                        if (cCliente > 0)
                            item.Client = row.Cell(cCliente).GetFormattedString().Trim();

                        if (cFactura > 0)
                            item.Folio = NormalizeFolio(row.Cell(cFactura).GetFormattedString().Trim());

                        if (String.IsNullOrWhiteSpace(item.Folio))
                            item.Folio = NormalizeFolio(ExtractFolio(all));

                        if (String.IsNullOrWhiteSpace(item.Client) && cConcepto > 0)
                            item.Client = ExtractClient(row.Cell(cConcepto).GetFormattedString());

                        if (cFecha > 0)
                        {
                            DateTime dt;
                            if (row.Cell(cFecha).TryGetValue<DateTime>(out dt))
                                item.Date = dt;
                        }

                        if (cIva4 > 0)
                            item.Iva4 = AbsDecimal(row.Cell(cIva4).Value);

                        if (cIva16 > 0)
                            item.Iva16 = AbsDecimal(row.Cell(cIva16).Value);

                        if (cIva4 == 0 && cIva16 == 0)
                        {
                            if (cDebe > 0)
                                item.Iva4 = AbsDecimal(row.Cell(cDebe).Value);
                            if (cHaber > 0)
                                item.Iva16 = AbsDecimal(row.Cell(cHaber).Value);
                        }

                        result.Add(item);
                    }
                }
            }

            List<AuxItem> consolidated = result
                .GroupBy(x => x.IncomeNo + "|" + x.InvoiceUuid,
                    StringComparer.OrdinalIgnoreCase)
                .Select(g =>
                {
                    AuxItem x = new AuxItem();
                    x.IncomeNo = g.First().IncomeNo;
                    x.Date = g.Select(v => v.Date).FirstOrDefault(v => v.HasValue);
                    x.InvoiceUuid = g.First().InvoiceUuid;
                    x.Client = g.Select(v => v.Client)
                        .FirstOrDefault(v => !String.IsNullOrWhiteSpace(v)) ?? "";
                    x.Folio = g.Select(v => v.Folio)
                        .FirstOrDefault(v => !String.IsNullOrWhiteSpace(v)) ?? "";
                    x.Iva4 = g.Sum(v => v.Iva4);
                    x.Iva16 = g.Sum(v => v.Iva16);
                    return x;
                })
                .ToList();

            return consolidated;
        }

        private static ComplementBook LoadComplements(string path)
        {
            ComplementBook output = new ComplementBook();

            using (XLWorkbook wb = new XLWorkbook(path))
            {
                foreach (IXLWorksheet ws in wb.Worksheets)
                {
                    IXLRange used = ws.RangeUsed();
                    if (used == null)
                        continue;

                    int headerRow = FindHeaderRow(ws, 30);
                    List<HeaderDef> headers = ws.Row(headerRow).CellsUsed()
                        .Select(c => new HeaderDef(
                            c.Address.ColumnNumber,
                            c.GetFormattedString().Trim()))
                        .Where(x => !String.IsNullOrWhiteSpace(x.Name))
                        .ToList();

                    foreach (HeaderDef hd in headers)
                    {
                        if (!output.Headers.Any(x =>
                            NormalizeHeader(x) == NormalizeHeader(hd.Name)))
                            output.Headers.Add(hd.Name);
                    }

                    foreach (IXLRow row in ws.RowsUsed())
                    {
                        if (row.RowNumber() <= headerRow)
                            continue;

                        Dictionary<string, string> fields =
                            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                        foreach (HeaderDef hd in headers)
                            fields[hd.Name] = row.Cell(hd.Column).GetFormattedString().Trim();

                        List<string> uuids = fields.Values
                            .SelectMany(v => Regex.Matches(v ?? "", UuidPattern)
                                .Cast<Match>()
                                .Select(m => m.Value.ToUpperInvariant()))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .ToList();

                        if (uuids.Count == 0)
                            continue;

                        ComplementRow cr = new ComplementRow();
                        cr.Fields = fields;

                        cr.InvoiceUuid = FindUuidByHeader(
                            fields,
                            new string[] { "DOCUMENT", "FACTURA", "RELACION" },
                            "");

                        cr.ComplementUuid = FindUuidByHeader(
                            fields,
                            new string[] { "COMPLEMENT", "PAGO", "FISCAL" },
                            cr.InvoiceUuid);

                        if (String.IsNullOrWhiteSpace(cr.InvoiceUuid))
                            cr.InvoiceUuid = uuids[uuids.Count - 1];

                        if (String.IsNullOrWhiteSpace(cr.ComplementUuid))
                            cr.ComplementUuid = uuids
                                .FirstOrDefault(x =>
                                    !x.Equals(cr.InvoiceUuid, StringComparison.OrdinalIgnoreCase)) ?? "";

                        cr.ImpPagado = FindDecimalByHeader(fields,
                            new string[] { "IMPPAGADO", "IMP PAGADO", "IMPORTE PAGADO" });

                        cr.PaymentAmount = FindDecimalByHeader(fields,
                            new string[] { "MONTO", "MONTO PAGO", "MONTO PAGADO" });

                        output.Rows.Add(cr);
                    }
                }
            }

            return output;
        }

        private static void LinkComplements(
            List<AuxItem> aux,
            ComplementBook comp)
        {
            Dictionary<string, List<ComplementRow>> byInvoice =
                comp.Rows
                    .Where(x => !String.IsNullOrWhiteSpace(x.InvoiceUuid))
                    .GroupBy(x => x.InvoiceUuid, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(g => g.Key, g => g.ToList(),
                        StringComparer.OrdinalIgnoreCase);

            foreach (AuxItem item in aux)
            {
                List<ComplementRow> rows;
                if (byInvoice.TryGetValue(item.InvoiceUuid, out rows))
                {
                    item.Complement = rows
                        .OrderByDescending(x => x.ImpPagado > 0)
                        .ThenByDescending(x => x.PaymentAmount > 0)
                        .FirstOrDefault();
                }
            }
        }

        private static void MatchDeposits(
            List<BankDeposit> deposits,
            List<AuxItem> aux)
        {
            List<IncomeGroup> groups = aux
                .Where(x => !String.IsNullOrWhiteSpace(x.IncomeNo))
                .GroupBy(x => x.IncomeNo, StringComparer.OrdinalIgnoreCase)
                .Select(g =>
                {
                    IncomeGroup x = new IncomeGroup();
                    x.IncomeNo = g.Key;
                    x.Items = g.ToList();
                    x.Client = g.Select(v => v.Client)
                        .FirstOrDefault(v => !String.IsNullOrWhiteSpace(v)) ?? "";
                    x.RelatedPaid = x.Items
                        .Where(v => v.Complement != null)
                        .Sum(v => v.Complement.ImpPagado);
                    return x;
                })
                .ToList();

            HashSet<string> used =
                new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (BankDeposit dep in deposits.OrderBy(x => x.Order))
            {
                IncomeGroup best = null;
                int bestScore = 0;
                int secondScore = 0;

                foreach (IncomeGroup g in groups.Where(x => !used.Contains(x.IncomeNo)))
                {
                    int score = Score(dep, g);

                    if (score > bestScore)
                    {
                        secondScore = bestScore;
                        bestScore = score;
                        best = g;
                    }
                    else if (score > secondScore)
                    {
                        secondScore = score;
                    }
                }

                if (best != null && bestScore >= 3 && bestScore >= secondScore + 1)
                {
                    dep.Income = best;
                    used.Add(best.IncomeNo);
                    AllocateCollection(dep, best);
                }
            }
        }

        private static int Score(BankDeposit dep, IncomeGroup g)
        {
            int score = 0;
            string concept = NormalizeText(dep.Concept);
            string client = NormalizeText(g.Client);

            if (!String.IsNullOrWhiteSpace(client))
            {
                if (concept.Contains(client))
                    score += 5;
                else
                {
                    List<string> words = client.Split(' ')
                        .Where(x => x.Length >= 4)
                        .Distinct()
                        .ToList();

                    int hits = words.Count(x => concept.Contains(x));
                    if (hits >= 2)
                        score += 3;
                    else if (hits == 1)
                        score += 1;
                }
            }

            foreach (string f in g.Items
                .Select(x => NormalizeFolio(x.Folio))
                .Where(x => x.Length > 0))
            {
                if (concept.Contains(f))
                    score += 4;
            }

            if (g.RelatedPaid > 0)
            {
                decimal diff = Math.Abs(dep.Amount - g.RelatedPaid);
                if (diff <= 0.05m)
                    score += 6;
                else if (diff <= 1m)
                    score += 5;
                else if (diff <= 800m)
                    score += 2;
            }

            List<DateTime> dates = g.Items
                .Where(x => x.Date.HasValue)
                .Select(x => x.Date.Value.Date)
                .ToList();

            if (dates.Count > 0)
            {
                int minDays = dates.Min(d =>
                    Math.Abs((dep.Date.Date - d).Days));

                if (minDays == 0)
                    score += 2;
                else if (minDays <= 3)
                    score += 1;
            }

            return score;
        }

        private static void AllocateCollection(BankDeposit dep, IncomeGroup group)
        {
            decimal assigned = 0;

            foreach (AuxItem item in group.Items)
            {
                decimal paid = item.Complement == null ? 0 : item.Complement.ImpPagado;
                if (paid <= 0)
                    continue;

                if (item.Iva16 > 0 && item.Iva4 <= 0)
                    item.Allocated16 = paid;
                else if (item.Iva4 > 0 && item.Iva16 <= 0)
                    item.Allocated4 = paid;
                else if (item.Iva4 > 0 && item.Iva16 > 0)
                {
                    decimal gross16 = item.Iva16 / 0.16m + item.Iva16;
                    decimal gross4 = item.Iva4 / 0.04m + item.Iva4;
                    decimal gross = gross16 + gross4;

                    if (gross > 0)
                    {
                        item.Allocated16 =
                            Math.Round(paid * gross16 / gross, 2);
                        item.Allocated4 = paid - item.Allocated16;
                    }
                }

                assigned += item.Allocated16 + item.Allocated4;
            }

            dep.Remainder = dep.Amount - assigned;
            if (dep.Remainder < 0 && Math.Abs(dep.Remainder) < 1m)
                dep.Remainder = 0;

            if (dep.Remainder < 0)
                dep.Remainder = 0;

            string c = dep.Concept.ToUpperInvariant();

            if (c.Contains("REEMBOLS"))
                dep.RemainderClass = "REEMBOLSABLE";
            else if (c.Contains("DEVOLU"))
                dep.RemainderClass = "DEVOLUCION";
            else if (c.Contains("TRASPAS"))
                dep.RemainderClass = "TRASPASO";
            else
                dep.RemainderClass = "NO IDENTIFICADO";
        }

        private void ExportExcel(string outPath)
        {
            string template =
                Path.Combine(AppContext.BaseDirectory, "cedulas INGRESOS(3).xlsx");

            if (!File.Exists(template))
                throw new FileNotFoundException(
                    "No encontré la plantilla cedulas INGRESOS(3).xlsx.");

            File.Copy(template, outPath, true);

            using (XLWorkbook wb = new XLWorkbook(outPath))
            {
                IXLWorksheet ws = wb.Worksheet("INGRESOS");

                string[] fixedHeaders =
                {
                    "INGRESO",
                    "Fecha",
                    "Concepto / Referencia",
                    "DEPOSITO DLL",
                    "DEPOSITO MN",
                    "COBRANZA tasa 16",
                    "COBRANZA tasa 4",
                    "cobranza 0",
                    "INTERESES",
                    "reembolsables",
                    "DEVOLUCION",
                    "NO IDENTIFICADO",
                    "TRASPASO",
                    "TRASPASO INVERSION",
                    "PRESTAMO",
                    "FACTURA",
                    "UUID FACTURA",
                    "UUID COMPLEMENTO DE PAGO",
                    "DIFERENCIA"
                };

                ws.Range("A4:ZZ2004").Clear(XLClearOptions.Contents);

                for (int i = 0; i < fixedHeaders.Length; i++)
                    ws.Cell(4, i + 1).Value = fixedHeaders[i];

                int compStartCol = 20;

                for (int i = 0; i < _complements.Headers.Count; i++)
                    ws.Cell(4, compStartCol + i).Value = _complements.Headers[i];

                int outRow = 5;

                foreach (BankDeposit dep in _deposits.OrderBy(x => x.Order))
                {
                    List<AuxItem> details = dep.Income == null
                        ? new List<AuxItem>()
                        : dep.Income.Items;

                    if (details.Count == 0)
                    {
                        WriteBankBase(ws, outRow, dep, "", true);

                        string up = dep.Concept.ToUpperInvariant();
                        if (up.Contains("REEMBOLS"))
                            ws.Cell(outRow, 10).Value = dep.Amount;
                        else if (up.Contains("DEVOLU"))
                            ws.Cell(outRow, 11).Value = dep.Amount;
                        else if (up.Contains("TRASPAS"))
                            ws.Cell(outRow, 13).Value = dep.Amount;
                        else
                            ws.Cell(outRow, 12).Value = dep.Amount;

                        ws.Cell(outRow, 19).Value = 0;
                        outRow++;
                        continue;
                    }

                    for (int i = 0; i < details.Count; i++)
                    {
                        AuxItem d = details[i];
                        bool first = (i == 0);

                        WriteBankBase(
                            ws,
                            outRow,
                            dep,
                            dep.Income.IncomeNo,
                            first);

                        if (d.Allocated16 != 0)
                            ws.Cell(outRow, 6).Value = d.Allocated16;

                        if (d.Allocated4 != 0)
                            ws.Cell(outRow, 7).Value = d.Allocated4;

                        if (first && dep.Remainder > 0)
                        {
                            if (dep.RemainderClass == "REEMBOLSABLE")
                                ws.Cell(outRow, 10).Value = dep.Remainder;
                            else if (dep.RemainderClass == "DEVOLUCION")
                                ws.Cell(outRow, 11).Value = dep.Remainder;
                            else if (dep.RemainderClass == "TRASPASO")
                                ws.Cell(outRow, 13).Value = dep.Remainder;
                            else
                                ws.Cell(outRow, 12).Value = dep.Remainder;
                        }

                        ws.Cell(outRow, 16).Value = d.Folio;
                        ws.Cell(outRow, 17).Value = d.InvoiceUuid;
                        ws.Cell(outRow, 18).Value =
                            d.Complement == null ? "" : d.Complement.ComplementUuid;

                        if (d.Complement != null)
                        {
                            for (int h = 0; h < _complements.Headers.Count; h++)
                            {
                                string header = _complements.Headers[h];
                                string value;

                                if (d.Complement.Fields.TryGetValue(header, out value))
                                    ws.Cell(outRow, compStartCol + h).Value = value;
                            }
                        }

                        if (first)
                        {
                            decimal classified =
                                details.Sum(x => x.Allocated16 + x.Allocated4)
                                + dep.Remainder;

                            ws.Cell(outRow, 19).Value =
                                dep.Amount - classified;
                        }

                        outRow++;
                    }
                }

                ws.SheetView.FreezeRows(4);
                ws.Row(4).Style.Font.Bold = true;
                ws.Row(4).Style.Alignment.WrapText = true;
                ws.Column(1).Width = 12;
                ws.Column(2).Width = 13;
                ws.Column(3).Width = 58;
                ws.Column(16).Width = 16;
                ws.Column(17).Width = 40;
                ws.Column(18).Width = 40;
                ws.Column(19).Width = 14;

                if (outRow > 5)
                {
                    ws.Range("B5:B" + (outRow - 1))
                        .Style.DateFormat.Format = "dd/mm/yyyy";

                    ws.Range("C5:C" + (outRow - 1))
                        .Style.Alignment.WrapText = true;
                }

                wb.Save();
            }
        }

        private static void WriteBankBase(
            IXLWorksheet ws,
            int row,
            BankDeposit dep,
            string incomeNo,
            bool firstDetail)
        {
            ws.Cell(row, 1).Value = incomeNo;
            ws.Cell(row, 2).Value = dep.Date;
            ws.Cell(row, 3).Value = dep.Concept;

            if (firstDetail)
                ws.Cell(row, 5).Value = dep.Amount;
        }

        private static int FindHeaderRow(IXLWorksheet ws, int maxRows)
        {
            int last = ws.LastRowUsed() == null
                ? 1
                : ws.LastRowUsed().RowNumber();

            int limit = Math.Min(maxRows, last);
            int bestRow = 1;
            int bestScore = -1;

            for (int r = 1; r <= limit; r++)
            {
                int score = ws.Row(r).CellsUsed()
                    .Count(c => c.GetFormattedString().Trim().Length > 0);

                if (score > bestScore)
                {
                    bestScore = score;
                    bestRow = r;
                }
            }

            return bestRow;
        }

        private static Dictionary<string, int> BuildHeaderMap(
            IXLWorksheet ws,
            int row)
        {
            Dictionary<string, int> d =
                new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach (IXLCell c in ws.Row(row).CellsUsed())
            {
                string h = NormalizeHeader(c.GetFormattedString());

                if (h.Length > 0 && !d.ContainsKey(h))
                    d[h] = c.Address.ColumnNumber;
            }

            return d;
        }

        private static int FindColumn(
            Dictionary<string, int> map,
            string[] aliases)
        {
            foreach (string alias in aliases)
            {
                string a = NormalizeHeader(alias);

                foreach (KeyValuePair<string, int> kv in map)
                {
                    if (kv.Key == a || kv.Key.Contains(a))
                        return kv.Value;
                }
            }

            return 0;
        }

        private static string NormalizeHeader(string s)
        {
            if (s == null)
                return "";

            string x = s.ToUpperInvariant().Trim();
            return Regex.Replace(x, @"[^A-ZÁÉÍÓÚÜÑ0-9]+", " ").Trim();
        }

        private static string NormalizeText(string s)
        {
            return Regex.Replace(
                NormalizeHeader(s),
                @"\s+",
                " ").Trim();
        }

        private const string UuidPattern =
            @"\b[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\b";

        private static string ExtractUuid(string text)
        {
            Match m = Regex.Match(text ?? "", UuidPattern);
            return m.Success ? m.Value : "";
        }

        private static string ExtractFolio(string text)
        {
            Match m = Regex.Match(
                text ?? "",
                @"\bF-\s*([A-Za-z0-9\-]+)\b",
                RegexOptions.IgnoreCase);

            return m.Success ? m.Groups[1].Value : "";
        }

        private static string NormalizeFolio(string value)
        {
            string x = value == null ? "" : value.Trim();
            x = Regex.Replace(
                x,
                @"^\s*F-\s*",
                "",
                RegexOptions.IgnoreCase);

            return x.Trim();
        }

        private static string ExtractClient(string concept)
        {
            string[] parts = (concept ?? "")
                .Split('|')
                .Select(x => x.Trim())
                .ToArray();

            return parts.Length >= 2 ? parts[1] : "";
        }

        private static decimal ParseDecimal(string value)
        {
            if (String.IsNullOrWhiteSpace(value))
                return 0;

            string x = value.Replace("$", "").Replace(" ", "");

            decimal d;
            if (Decimal.TryParse(
                x,
                NumberStyles.Any,
                CultureInfo.GetCultureInfo("es-MX"),
                out d))
                return d;

            if (Decimal.TryParse(
                x,
                NumberStyles.Any,
                CultureInfo.GetCultureInfo("en-US"),
                out d))
                return d;

            return 0;
        }

        private static decimal AbsDecimal(XLCellValue value)
        {
            if (value.IsNumber)
                return Math.Abs((decimal)value.GetNumber());

            return Math.Abs(ParseDecimal(value.ToString()));
        }

        private static string FindUuidByHeader(
            Dictionary<string, string> fields,
            string[] words,
            string exclude)
        {
            foreach (KeyValuePair<string, string> kv in fields)
            {
                string h = NormalizeHeader(kv.Key);
                bool ok = words.Any(w =>
                    h.Contains(NormalizeHeader(w)));

                if (!ok)
                    continue;

                Match m = Regex.Match(kv.Value ?? "", UuidPattern);

                if (m.Success &&
                    !m.Value.Equals(exclude, StringComparison.OrdinalIgnoreCase))
                    return m.Value.ToUpperInvariant();
            }

            return "";
        }

        private static decimal FindDecimalByHeader(
            Dictionary<string, string> fields,
            string[] words)
        {
            foreach (KeyValuePair<string, string> kv in fields)
            {
                string h = NormalizeHeader(kv.Key);

                if (words.Any(w => h.Contains(NormalizeHeader(w))))
                {
                    decimal d = ParseDecimal(kv.Value);
                    if (d != 0)
                        return d;
                }
            }

            return 0;
        }

        private sealed class HeaderDef
        {
            public int Column;
            public string Name;

            public HeaderDef(int column, string name)
            {
                Column = column;
                Name = name;
            }
        }

        private sealed class ComplementBook
        {
            public List<string> Headers = new List<string>();
            public List<ComplementRow> Rows = new List<ComplementRow>();
        }

        private sealed class ComplementRow
        {
            public string InvoiceUuid = "";
            public string ComplementUuid = "";
            public decimal ImpPagado = 0;
            public decimal PaymentAmount = 0;
            public Dictionary<string, string> Fields =
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        private sealed class AuxItem
        {
            public string IncomeNo = "";
            public DateTime? Date = null;
            public string InvoiceUuid = "";
            public string Client = "";
            public string Folio = "";
            public decimal Iva4 = 0;
            public decimal Iva16 = 0;
            public ComplementRow Complement = null;
            public decimal Allocated4 = 0;
            public decimal Allocated16 = 0;
        }

        private sealed class IncomeGroup
        {
            public string IncomeNo = "";
            public string Client = "";
            public List<AuxItem> Items = new List<AuxItem>();
            public decimal RelatedPaid = 0;
        }

        private sealed class BankDeposit
        {
            public int Order = 0;
            public DateTime Date;
            public decimal Amount = 0;
            public string Concept = "";
            public IncomeGroup Income = null;
            public decimal Remainder = 0;
            public string RemainderClass = "";
        }
    }
}
