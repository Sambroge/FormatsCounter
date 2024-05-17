using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TFlex.Model;
using TFlex.Model.Model2D;
using TFlex.Model.Model3D;

namespace NewMacroNamespace
{
    public class NewMacroClass
    {
        public static void NewMacro()
        {
            Document document = TFlex.Application.ActiveDocument;
            document.BeginChanges("Создание переменных");
            var pages = document.GetPages();
            double number = 0;
            List<string> formats = new List<string>() { "A4", "A3", "A2", "A1", "A0", "A4x3", "A4x4", "A4x5", "A4x6", "A4x7",
            "A4x8", "A4x9","A3x3", "A3x4","A3x5", "A3x6","A3x7", "A2x3", "A2x4", "A2x5", "A1x3", "A1x4", "A0x2", "A0x3"};
            Dictionary<string, double> counter = new Dictionary<string, double>();
            foreach (string format in formats)
            {
                if (document.FindVariable(format) == null) { Variable a = new Variable(document, format, number, false); }
                counter.Add(format, number);
            }
            foreach (var b in pages)
            {
                if (b.PageType == PageType.Normal)
                {
                    if (counter.ContainsKey(b.Properties.Paper.Format))
                    {
                        counter[b.Properties.Paper.Format]++;
                    }
                }
            }

            foreach (var variable in counter)
            {
                document.FindVariable(variable.Key).RealValue = variable.Value;
            }
            document.EndChanges();
        }

    }
}
