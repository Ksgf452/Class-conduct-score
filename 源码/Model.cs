using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace CaoXing
{
    /// <summary>一条操行记录。</summary>
    public class RecordItem
    {
        public int AspectIndex;          // 1 开始的方面序号，0 表示模板中未标注方面
        public string AspectLabel;       // “方面一”等，可能为空
        public string Item;              // 事项
        public decimal Delta;            // 加减分
        public string Raw;               // 原始文本片段（含括号分数），用于无损回写
        public bool Imported;            // 是否从已有表格解析而来
        public DateTime? Time;           // 本次程序内添加的时间；模板解析的记录为空

        public string DeltaText
        {
            get { return (Delta > 0 ? "+" : "") + Format.Decimal(Delta); }
        }
    }

    /// <summary>一个模块的数据。</summary>
    public class ModuleData
    {
        public decimal Score;            // 当前模块分
        public decimal Adjustment;       // 无法解析的差额（通常为 0）
        public bool HasScore;            // 原表格该模块是否有分数
        public List<RecordItem> Records = new List<RecordItem>();
        public bool Dirty;               // 是否需要在保存时重写该模块
        public string OriginalText = ""; // 读取时的原始明细

        public decimal SumDelta()
        {
            decimal sum = 0m;
            for (int i = 0; i < Records.Count; i++) sum += Records[i].Delta;
            return sum;
        }
    }

    /// <summary>模块定义：名称、方面数量。</summary>
    public class ModuleDef
    {
        public string Name;
        public int AspectCount;

        public ModuleDef() { }
        public ModuleDef(string name, int aspectCount)
        {
            Name = name;
            AspectCount = aspectCount;
        }

        public ModuleDef Clone()
        {
            return new ModuleDef(Name, AspectCount);
        }
    }

    /// <summary>一名学生。</summary>
    public class Student
    {
        public int Seq;
        public string No = "";
        public string Name = "";
        public int RowIndex;             // 在 Excel 中的行号，0 表示新学生
        public List<ModuleData> Modules = new List<ModuleData>();

        public string Display
        {
            get
            {
                if (string.IsNullOrEmpty(No)) return Name;
                return No + "  " + Name;
            }
        }
    }

    public class StudentSeed
    {
        public string No;
        public string Name;
        public StudentSeed() { }
        public StudentSeed(string no, string name) { No = no; Name = name; }
    }

    /// <summary>整个班级的数据与计算规则。</summary>
    public class ClassData
    {
        public const string AppVersion = "1.1.0";

        public string ClassName = "";
        public decimal InitialScore = 80m;
        public List<ModuleDef> Modules = new List<ModuleDef>();
        public List<Student> Students = new List<Student>();

        public string SourceFile = "";
        public bool FromTemplate;
        public bool RosterDirty;

        public int DataStartRow = 3;
        public int AverageColumn;
        public int[] ScoreColumns;
        public int[] DetailColumns;

        public List<string> Warnings = new List<string>();

        public int ModuleCount { get { return Modules.Count; } }

        public decimal GetModuleScore(Student s, int moduleIndex)
        {
            ModuleData m = s.Modules[moduleIndex];
            return InitialScore + m.SumDelta() + m.Adjustment;
        }

        public decimal GetAverage(Student s)
        {
            if (Modules.Count == 0) return 0m;
            decimal sum = 0m;
            for (int i = 0; i < Modules.Count; i++) sum += GetModuleScore(s, i);
            return sum / Modules.Count;
        }

        public decimal GetModuleSum(Student s)
        {
            decimal sum = 0m;
            for (int i = 0; i < Modules.Count; i++) sum += GetModuleScore(s, i);
            return sum;
        }

        public static string AspectLabel(int index)
        {
            if (index <= 0) return "";
            return "方面" + Cn.Number(index);
        }

        public Student NewStudent(string no, string name)
        {
            Student s = new Student();
            s.Seq = Students.Count + 1;
            s.No = (no ?? "").Trim();
            s.Name = (name ?? "").Trim();
            for (int i = 0; i < Modules.Count; i++)
            {
                ModuleData m = new ModuleData();
                m.Score = InitialScore;
                m.HasScore = true;
                m.Dirty = true;
                s.Modules.Add(m);
            }
            Students.Add(s);
            return s;
        }

        /// <summary>添加一条记录，返回记录本身。</summary>
        public RecordItem AddRecord(Student s, int moduleIndex, int aspectIndex, string item, decimal delta, DateTime time)
        {
            ModuleData m = s.Modules[moduleIndex];
            string label = AspectLabel(aspectIndex);
            string raw = label + (item ?? "").Trim() + "（" + Format.Decimal(delta) + "分）";
            RecordItem r = new RecordItem();
            r.AspectIndex = aspectIndex;
            r.AspectLabel = label;
            r.Item = (item ?? "").Trim();
            r.Delta = delta;
            r.Raw = raw;
            r.Imported = false;
            r.Time = time;
            m.Records.Add(r);
            m.Dirty = true;
            m.Score = GetModuleScore(s, moduleIndex);
            return r;
        }

        public void Recalculate(Student s, int moduleIndex)
        {
            ModuleData m = s.Modules[moduleIndex];
            m.Score = GetModuleScore(s, moduleIndex);
            m.Dirty = true;
        }

        public string DetailText(ModuleData m)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < m.Records.Count; i++) sb.Append(m.Records[i].Raw ?? "");
            return sb.ToString();
        }

        public string ModuleName(int index)
        {
            if (index < 0 || index >= Modules.Count) return "";
            return Modules[index].Name;
        }

        public ClassData CloneConfig()
        {
            ClassData c = new ClassData();
            c.ClassName = ClassName;
            c.InitialScore = InitialScore;
            for (int i = 0; i < Modules.Count; i++) c.Modules.Add(Modules[i].Clone());
            return c;
        }

        public string RosterSignature()
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < Students.Count; i++)
            {
                sb.Append(Students[i].No).Append('\u0001').Append(Students[i].Name).Append('\u0002');
            }
            return sb.ToString();
        }
    }

    /// <summary>数字格式化帮助。</summary>
    public static class Format
    {
        public static string Decimal(decimal d)
        {
            return d.ToString("0.###", CultureInfo.InvariantCulture);
        }

        public static string Decimal2(decimal d)
        {
            return d.ToString("0.##", CultureInfo.InvariantCulture);
        }

        public static string Raw(decimal d)
        {
            return d.ToString(CultureInfo.InvariantCulture);
        }

        public static decimal ParseDecimal(string text, decimal fallback)
        {
            if (string.IsNullOrEmpty(text)) return fallback;
            decimal d;
            if (decimal.TryParse(text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return d;
            if (decimal.TryParse(text.Trim(), NumberStyles.Float, CultureInfo.CurrentCulture, out d)) return d;
            return fallback;
        }

        public static string ScoreWithSign(decimal d)
        {
            return (d > 0 ? "+" : "") + Decimal(d);
        }
    }
}


