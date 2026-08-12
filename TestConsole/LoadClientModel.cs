using System;
using System.Collections.Generic;
using System.Linq;
using MyOlap.Core;
using MyOlap.Data;

/// <summary>
/// Rebuilds the client's Sales Analysis model from scratch.
/// Safe to re-run — skips anything already present.
/// </summary>
static class LoadClientModel
{
    static SqliteRepository repo = null!;

    // Member cache: (dimId, name) → memberId
    static readonly Dictionary<(long, string), long> _cache =
        new Dictionary<(long, string), long>(
            EqualityComparer<(long, string)>.Create(
                (a, b) => a.Item1 == b.Item1 && string.Equals(a.Item2, b.Item2, StringComparison.OrdinalIgnoreCase),
                x => HashCode.Combine(x.Item1, x.Item2.ToLowerInvariant())));

    static void CacheDim(long dimId)
    {
        foreach (var m in repo.GetMembers(dimId))
            _cache.TryAdd((dimId, m.Name), m.Id);
    }

    static long M(long dimId, string name, long? parent = null, string consol = "+",
                  string? timeBal = null, string? formula = null)
    {
        if (_cache.TryGetValue((dimId, name), out var existing)) return existing;
        int sort = parent.HasValue ? repo.GetChildren(parent.Value).Count : repo.GetRootMembers(dimId).Count;
        var id = repo.InsertMember(new Member
        {
            DimensionId    = dimId,
            ParentId       = parent,
            Name           = name,
            Level          = parent.HasValue ? 1 : 0,
            SortOrder      = sort,
            ConsolOperator = consol,
            TimeBalance    = timeBal ?? "",
            Formula        = formula ?? ""
        });
        _cache[(dimId, name)] = id;
        return id;
    }

    static long D(long modelId, string name, DimensionType type, ref int sort, List<Dimension> dims)
    {
        var existing = dims.FirstOrDefault(d => d.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        if (existing != null) { CacheDim(existing.Id); return existing.Id; }
        var id = repo.InsertDimension(new Dimension { ModelId = modelId, Name = name, DimType = type, SortOrder = sort++ });
        return id;
    }

    public static void Run()
    {
        repo = SqliteRepository.Instance;
        repo.EnsureDatabaseCreated();
        _cache.Clear();

        // ── Get or create the model ──────────────────────────────────────────
        var models = repo.GetAllModels();
        OlapModel model;
        if (models.Count == 0)
        {
            var newId = repo.InsertModel(new OlapModel { Name = "Sales Analysis", Description = "Client sales model (loaded from TestData)", CreatedUtc = DateTime.UtcNow });
            model = repo.GetAllModels().First(m => m.Id == newId);
            Console.WriteLine($"  [NEW] Created model Id={model.Id} '{model.Name}'");
        }
        else
        {
            model = models.OrderBy(m => m.Id).First();
            Console.WriteLine($"  Model: Id={model.Id}  Name='{model.Name}'");
        }

        var dims = repo.GetDimensions(model.Id);
        int nextSort = dims.Any() ? dims.Max(d => d.SortOrder) + 1 : 0;

        // ── MEASURE ─────────────────────────────────────────────────────────
        long measId = D(model.Id, "Measure", DimensionType.Measure, ref nextSort, dims);
        dims = repo.GetDimensions(model.Id);
        CacheDim(measId);
        if (repo.GetRootMembers(measId).Count == 0)
        {
            // Top-level hierarchy matching Sales_Dimensions.xlsx
            long tp  = M(measId, "Trading Profit");
            long sal = M(measId, "Sales",             tp,  "+", "Flow");
            long cos = M(measId, "Cost of Sales",     tp,  "-", "Flow");
            // Sales sub-hierarchy
            long sg  = M(measId, "Sales - Goods",     sal, "+");
            long sgs = M(measId, "Sales - Goods - Stock", sg, "+");
            M(measId, "600000", sgs, "+"); M(measId, "600010", sgs, "+");
            M(measId, "600020", sgs, "+"); M(measId, "600030", sgs, "+");
            M(measId, "600040", sgs, "+"); M(measId, "600050", sgs, "+");
            long sgsp = M(measId, "Sales - Goods - Specials / Non-Stock", sg, "+");
            M(measId, "600100", sgsp, "+"); M(measId, "600110", sgsp, "+");
            M(measId, "600120", sgsp, "+"); M(measId, "600130", sgsp, "+");
            long sgmi = M(measId, "Sales - Goods - Manufactured / Installed", sg, "+");
            M(measId, "600200", sgmi, "+"); M(measId, "600210", sgmi, "+");
            M(measId, "600220", sgmi, "+"); M(measId, "600230", sgmi, "+");
            long sgdp = M(measId, "Sales - Delivery & Packaging Income", sg, "+");
            M(measId, "600300", sgdp, "+");

            long sd  = M(measId, "Sales - Directs",   sal, "+");
            M(measId, "601000", sd, "+"); M(measId, "601010", sd, "+");

            long ser = M(measId, "Sales - External - Returns & Allowances", sal, "+");
            long sra = M(measId, "Sales - Returns & Allowances", ser, "+");
            M(measId, "602000", sra, "+");

            long srd = M(measId, "Sales - Rebates & Discounts", sal, "+");
            long srb = M(measId, "Sales - Rebates", srd, "+"); M(measId, "603000", srb, "+");
            long sdc = M(measId, "Sales - Discounts", srd, "+");
            M(measId, "603100", sdc, "+"); M(measId, "603110", sdc, "+");

            long sng = M(measId, "Sales - Non-Goods", sal, "+");
            long shr = M(measId, "Sales - Services - Hire / Rental", sng, "+"); M(measId, "604000", shr, "+");
            long sms = M(measId, "Sales - Services - Maintenance / After Sales Services", sng, "+"); M(measId, "604100", sms, "+");
            long sof = M(measId, "Sales - Other Fees & Recharges", sng, "+");
            M(measId, "604200", sof, "+"); M(measId, "604210", sof, "+"); M(measId, "604220", sof, "+");
            long snc = M(measId, "Sales - Non-Goods - Non-Core Items", sng, "+"); M(measId, "604300", snc, "+");

            long si  = M(measId, "Sales - Internal", sal, "+");
            M(measId, "610000", si, "+"); M(measId, "610010", si, "+");

            // Cost of Sales sub-hierarchy
            long cgs = M(measId, "Cost of Goods Sold", cos, "+");
            long cgss= M(measId, "Cost of Goods Sold - Stock", cgs, "+");
            M(measId, "700000", cgss, "+"); M(measId, "700010", cgss, "+"); M(measId, "700020", cgss, "+");
            long cgsp= M(measId, "Cost of Goods Sold - Specials", cgs, "+");
            M(measId, "700100", cgsp, "+"); M(measId, "700110", cgsp, "+");
            long cgmi= M(measId, "Cost of Goods Sold - Manufactured / Installed", cgs, "+");
            M(measId, "700200", cgmi, "+"); M(measId, "700210", cgmi, "+"); M(measId, "700220", cgmi, "+");
            M(measId, "700230", cgmi, "+"); M(measId, "700240", cgmi, "+"); M(measId, "700250", cgmi, "+");
            long idc = M(measId, "Inbound Delivery Charges", cgs, "+");
            M(measId, "700300.126", idc, "+"); M(measId, "700310", idc, "+"); M(measId, "700320", idc, "+");

            long csd = M(measId, "Cost of Sales - Directs", cos, "+");
            M(measId, "701000", csd, "+");

            long iga = M(measId, "Inventory & Gross Profit Adjustments", cos, "+");
            long pia = M(measId, "Physical Inventory Adjustments", iga, "+");
            M(measId, "702000", pia, "+"); M(measId, "702010", pia, "+"); M(measId, "702020", pia, "+");
            M(measId, "702030", pia, "+"); M(measId, "702040", pia, "+");
            long gpa = M(measId, "Gross Profit Adjustments", iga, "+");
            M(measId, "702100", gpa, "+"); M(measId, "702110", gpa, "+"); M(measId, "702120", gpa, "+");
            M(measId, "702130", gpa, "+"); M(measId, "702140", gpa, "+"); M(measId, "702150", gpa, "+");
            M(measId, "702155.000", gpa, "+"); M(measId, "702160", gpa, "+"); M(measId, "702170", gpa, "+");
            M(measId, "702170.196", gpa, "+");

            long reb = M(measId, "Rebates / Discounts", cos, "+");
            long rbb = M(measId, "Rebates", reb, "+");
            M(measId, "703000", rbb, "+"); M(measId, "703010", rbb, "+"); M(measId, "703020", rbb, "+"); M(measId, "703030", rbb, "+");
            long dsc = M(measId, "Discounts", reb, "+");
            M(measId, "703100", dsc, "+"); M(measId, "703110", dsc, "+");

            long cng = M(measId, "Cost of Sales - Non-Goods", cos, "+");
            long css = M(measId, "Cost of Sales - Services", cng, "+");
            M(measId, "704000", css, "+"); M(measId, "704010", css, "+");

            long cin = M(measId, "Cost of Sales - Internal", cos, "+");
            M(measId, "710000", cin, "+"); M(measId, "710010", cin, "+"); M(measId, "710020", cin, "+");
            M(measId, "710030", cin, "+"); M(measId, "710040", cin, "+");
            M(measId, "910040", cin, "+"); M(measId, "910050", cin, "+");

            // Trading Metrics (second root)
            long tm = M(measId, "Trading Metrics");
            M(measId, "Margin", tm, "x", "Equation", "\"Trading Profit\" / \"Sales\" * 100");

            Console.WriteLine($"  [OK] Measure — {repo.GetMembers(measId).Count} members loaded");
        }
        else Console.WriteLine($"  [SKIP] Measure already has {repo.GetMembers(measId).Count} members");

        // ── TIME ─────────────────────────────────────────────────────────────
        long timeId = D(model.Id, "Time", DimensionType.Time, ref nextSort, dims);
        dims = repo.GetDimensions(model.Id);
        CacheDim(timeId);
        if (repo.GetRootMembers(timeId).Count == 0)
        {
            long tRoot  = M(timeId, "Time");
            long tOpBal = M(timeId, "OpBal",    tRoot, "x");
            long tYrTot = M(timeId, "YearTotal",tRoot, "+");
            long tYTD   = M(timeId, "YTD",      tRoot, "x");
            long tQ1 = M(timeId, "Q1", tYrTot, "+"); long tQ2 = M(timeId, "Q2", tYrTot, "+");
            long tQ3 = M(timeId, "Q3", tYrTot, "+"); long tQ4 = M(timeId, "Q4", tYrTot, "+");
            M(timeId, "Jan", tQ1, "+"); M(timeId, "Feb", tQ1, "+"); M(timeId, "Mar", tQ1, "+");
            M(timeId, "Apr", tQ2, "+"); M(timeId, "May", tQ2, "+"); M(timeId, "Jun", tQ2, "+");
            M(timeId, "Jul", tQ3, "+"); M(timeId, "Aug", tQ3, "+"); M(timeId, "Sep", tQ3, "+");
            M(timeId, "Oct", tQ4, "+"); M(timeId, "Nov", tQ4, "+"); M(timeId, "Dec", tQ4, "+");
            M(timeId, "YTDJan",  tYTD, "x", null, "\"Jan\"");
            M(timeId, "YTDFeb",  tYTD, "x", null, "\"Jan\" + \"Feb\"");
            M(timeId, "YTDMar",  tYTD, "x", null, "\"Jan\" + \"Feb\" + \"Mar\"");
            M(timeId, "YTDQtr1", tYTD, "x", null, "\"Q1\"");
            M(timeId, "YTDApr",  tYTD, "x", null, "\"Q1\" + \"Apr\"");
            M(timeId, "YTDMay",  tYTD, "x", null, "\"Q1\" + \"Apr\" + \"May\"");
            M(timeId, "YTDQtr2", tYTD, "x", null, "\"Q1\" + \"Q2\"");
            M(timeId, "YTDJul",  tYTD, "x", null, "\"Q1\" + \"Q2\" + \"Jul\"");
            M(timeId, "YTDAug",  tYTD, "x", null, "\"Q1\" + \"Q2\" + \"Jul\" + \"Aug\"");
            M(timeId, "YTDQtr3", tYTD, "x", null, "\"Q1\" + \"Q2\" + \"Q3\"");
            M(timeId, "YTDOct",  tYTD, "x", null, "\"Q1\" + \"Q2\" + \"Q3\" + \"Oct\"");
            M(timeId, "YTDNov",  tYTD, "x", null, "\"Q1\" + \"Q2\" + \"Q3\" + \"Oct\" + \"Nov\"");
            M(timeId, "YTDDec",  tYTD, "x", null, "\"Q1\" + \"Q2\" + \"Q3\" + \"Q4\"");
            Console.WriteLine($"  [OK] Time — {repo.GetMembers(timeId).Count} members");
        }
        else Console.WriteLine($"  [SKIP] Time already loaded");

        // ── YEAR ─────────────────────────────────────────────────────────────
        long yearId = D(model.Id, "Year", DimensionType.Year, ref nextSort, dims);
        dims = repo.GetDimensions(model.Id);
        CacheDim(yearId);
        if (repo.GetRootMembers(yearId).Count == 0)
        {
            long yAll = M(yearId, "All Years");
            M(yearId, "FY25", yAll, "+"); M(yearId, "FY26", yAll, "+"); M(yearId, "FY27", yAll, "+");
            Console.WriteLine($"  [OK] Year — FY25/FY26/FY27");
        }
        else Console.WriteLine($"  [SKIP] Year already loaded");

        // ── VIEW ─────────────────────────────────────────────────────────────
        long viewId = D(model.Id, "View", DimensionType.View, ref nextSort, dims);
        dims = repo.GetDimensions(model.Id);
        CacheDim(viewId);
        if (repo.GetRootMembers(viewId).Count == 0)
        {
            long vAll = M(viewId, "All Views");
            M(viewId, "Actual", vAll, "+"); M(viewId, "Budget", vAll, "+"); M(viewId, "Forecast", vAll, "+");
            Console.WriteLine($"  [OK] View — Actual/Budget/Forecast");
        }
        else Console.WriteLine($"  [SKIP] View already loaded");

        // ── VERSION ──────────────────────────────────────────────────────────
        long verId = D(model.Id, "Version", DimensionType.Version, ref nextSort, dims);
        dims = repo.GetDimensions(model.Id);
        CacheDim(verId);
        if (repo.GetRootMembers(verId).Count == 0)
        {
            long veAll = M(verId, "All Versions");
            M(verId, "Draft", veAll, "+"); M(verId, "Final", veAll, "+"); M(verId, "What-If", veAll, "+");
            Console.WriteLine($"  [OK] Version — Draft/Final/What-If");
        }
        else Console.WriteLine($"  [SKIP] Version already loaded");

        // ── PRODUCT ──────────────────────────────────────────────────────────
        dims = repo.GetDimensions(model.Id);
        nextSort = dims.Any() ? dims.Max(d => d.SortOrder) + 1 : 5;
        long prodId = D(model.Id, "Product", DimensionType.UserDefined, ref nextSort, dims);
        dims = repo.GetDimensions(model.Id);
        CacheDim(prodId);
        if (repo.GetRootMembers(prodId).Count == 0)
        {
            long pAll = M(prodId, "All Products");
            foreach (var c in new[] { "FPROT","HVACR","MISC_PROD","PVF","PLUMBING","WATERWORKS","HYDRONICS","NOPROD" })
                M(prodId, c, pAll, "+");
            Console.WriteLine($"  [OK] Product — 8 product types");
        }
        else Console.WriteLine($"  [SKIP] Product already loaded");

        // ── CUSTOMER ─────────────────────────────────────────────────────────
        dims = repo.GetDimensions(model.Id);
        nextSort = dims.Max(d => d.SortOrder) + 1;
        long custId = D(model.Id, "Customer", DimensionType.UserDefined, ref nextSort, dims);
        dims = repo.GetDimensions(model.Id);
        CacheDim(custId);
        if (repo.GetRootMembers(custId).Count == 0)
        {
            long cAll = M(custId, "All Customers");
            foreach (var c in new[] { "COMML","INDL","INST","MISC_CUST","RESELLER","RESID","RETAIL","DIRECT_CUST" })
                M(custId, c, cAll, "+");
            Console.WriteLine($"  [OK] Customer — 8 segments");
        }
        else Console.WriteLine($"  [SKIP] Customer already loaded");

        // ── BU ───────────────────────────────────────────────────────────────
        dims = repo.GetDimensions(model.Id);
        nextSort = dims.Max(d => d.SortOrder) + 1;
        long buId = D(model.Id, "BU", DimensionType.UserDefined, ref nextSort, dims);
        dims = repo.GetDimensions(model.Id);
        CacheDim(buId);
        if (repo.GetRootMembers(buId).Count == 0)
        {
            long bAll  = M(buId, "All Stores");
            long bEast = M(buId, "EAST",    bAll, "+");
            long bCent = M(buId, "CENTRAL", bAll, "+");
            long bWest = M(buId, "WEST",    bAll, "+");
            Console.WriteLine($"  [OK] BU — EAST/CENTRAL/WEST (load stores via UI from Sales_Dimensions.xlsx)");
        }
        else Console.WriteLine($"  [SKIP] BU already loaded");

        // ── SAMPLE FACTS ─────────────────────────────────────────────────────
        dims = repo.GetDimensions(model.Id);
        var existingFacts = repo.GetAllFacts(model.Id);
        if (existingFacts.Count == 0)
        {
            var dimOrder = dims.OrderBy(d => d.SortOrder).Select(d => d.Id).ToList();

            long m600000 = _cache[(measId, "600000")];
            long m700000 = _cache[(measId, "700000")];
            long[] months = {
                _cache[(timeId,"Jan")], _cache[(timeId,"Feb")], _cache[(timeId,"Mar")],
                _cache[(timeId,"Apr")], _cache[(timeId,"May")], _cache[(timeId,"Jun")],
                _cache[(timeId,"Jul")], _cache[(timeId,"Aug")], _cache[(timeId,"Sep")],
                _cache[(timeId,"Oct")], _cache[(timeId,"Nov")], _cache[(timeId,"Dec")]
            };
            long yFY26   = _cache[(yearId, "FY26")];
            long vActual = _cache[(viewId, "Actual")];
            long vBudget = _cache[(viewId, "Budget")];
            long vnDraft = _cache[(verId,  "Draft")];
            long pFPROT  = _cache[(prodId, "FPROT")];
            long cCOMMl  = _cache[(custId, "COMML")];
            long bEAST   = _cache[(buId,   "EAST")];
            long bCENT   = _cache[(buId,   "CENTRAL")];

            string Key(long meas, long time, long year, long view, long ver, long prod, long cust, long bu) =>
                OlapEngine.BuildMemberKey(dimOrder, new Dictionary<long, long> {
                    { measId, meas }, { timeId, time }, { yearId, year },
                    { viewId, view }, { verId, ver }, { prodId, prod },
                    { custId, cust }, { buId, bu } });

            decimal[] actualSales = { 100000, 110000, 120000, 90000, 105000, 115000,
                                       95000, 108000, 112000, 98000, 107000, 125000 };
            decimal[] budgetSales = { 95000, 105000, 115000, 95000, 100000, 110000,
                                       100000, 105000, 110000, 102000, 110000, 120000 };

            var facts = new List<FactData>();
            for (int i = 0; i < 12; i++)
            {
                long mo = months[i];
                // Actual – EAST
                facts.Add(new FactData { ModelId = model.Id, MemberKey = Key(m600000, mo, yFY26, vActual, vnDraft, pFPROT, cCOMMl, bEAST), NumericValue = actualSales[i] });
                facts.Add(new FactData { ModelId = model.Id, MemberKey = Key(m700000, mo, yFY26, vActual, vnDraft, pFPROT, cCOMMl, bEAST), NumericValue = Math.Round(actualSales[i] * 0.60m) });
                // Actual – CENTRAL
                facts.Add(new FactData { ModelId = model.Id, MemberKey = Key(m600000, mo, yFY26, vActual, vnDraft, pFPROT, cCOMMl, bCENT), NumericValue = Math.Round(actualSales[i] * 0.80m) });
                facts.Add(new FactData { ModelId = model.Id, MemberKey = Key(m700000, mo, yFY26, vActual, vnDraft, pFPROT, cCOMMl, bCENT), NumericValue = Math.Round(actualSales[i] * 0.45m) });
                // Budget – EAST
                facts.Add(new FactData { ModelId = model.Id, MemberKey = Key(m600000, mo, yFY26, vBudget, vnDraft, pFPROT, cCOMMl, bEAST), NumericValue = budgetSales[i] });
                facts.Add(new FactData { ModelId = model.Id, MemberKey = Key(m700000, mo, yFY26, vBudget, vnDraft, pFPROT, cCOMMl, bEAST), NumericValue = Math.Round(budgetSales[i] * 0.60m) });
            }
            repo.InsertFactBatch(model.Id, facts);
            Console.WriteLine($"  [OK] Sample facts: {facts.Count} records (GL 600000=Sales, 700000=COGS | FY26 Actual+Budget / FPROT / COMML / EAST+CENTRAL)");
        }
        else Console.WriteLine($"  [SKIP] Facts already present ({existingFacts.Count} records)");

        // ── SUMMARY ──────────────────────────────────────────────────────────
        Console.WriteLine();
        Console.WriteLine("  Final state:");
        dims = repo.GetDimensions(model.Id);
        foreach (var d in dims.OrderBy(x => x.SortOrder))
            Console.WriteLine($"    [{d.SortOrder}] {d.Name,-12} ({d.DimType,-12}) — {repo.GetMembers(d.Id).Count,3} members");
        Console.WriteLine($"    Facts: {repo.GetAllFacts(model.Id).Count}");
    }
}
