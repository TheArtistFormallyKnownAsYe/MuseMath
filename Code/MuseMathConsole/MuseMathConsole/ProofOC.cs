using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Configuration;

namespace MuseMathConsole
{

    //I like enums, but this is not very helpful and will be deleted soon
    public enum Operation
    {
        Addition, Subtract, Multiply, Divide, Shape, Fraction
    }

    //Rename before release
    class ProofOC
    {
        public static int[] randomCapped4500;
        public static int[] randomCapped12;
        public static int[] randomDenominator;
        public static List<string> measurementItems;
        public static Random rnd;
        public static int workSheets;
        public static string sourcePath;
        public static string targetPath;

        private static readonly string[] Days =
            { "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday" };

        // Static constructor for initialization
        static ProofOC()
        {
            rnd = new Random();

            // Load settings from App.config
            sourcePath = ConfigurationManager.AppSettings["SourcePath"];
            targetPath = ConfigurationManager.AppSettings["TargetPath"];
            workSheets = int.TryParse(ConfigurationManager.AppSettings["WorkSheets"], out int ws) ? ws : 5;

            randomCapped4500 = new int[100];
            randomCapped12 = new int[100];
            randomDenominator = new int[] { 2, 3, 4, 5, 6, 8, 10, 12 };
            measurementItems = new List<string>
                    {
                        "bed", "chair", "desk", "table", "sofa", "pencil", "soda can", "swimming pool",
                        "paper clip", "grape", "leaf", "milk jug", "water bottle", "turtle", "horse",
                        "fish tank", "cup of water", "bed", "thumbtack", "staples", "grains of rice",
                        "stick of gum", "one raisin", "water jug", "lamp", "alarm clock",
                        "toothbrush", "cup", "washing machine", "mirror", "towel", "computer",
                        "medicine dropper", "teaspoon of oil", "small hand sanitizer", "ketchup packet",
                        "dollar bill", "elephant"
                    };

            for (int i = 0; i < 100; i++)
            {
                randomCapped4500[i] = rnd.Next(1, 4500);
                randomCapped12[i] = rnd.Next(1, 13);
            }
        }

        static void Main(string[] args)
        {
            InitializeConsole();

            // Arrays instantiation
            var addCalculations = CreateCalculationArray();
            var subCalculations = CreateCalculationArray();
            var mulCalculations = CreateCalculationArray();
            var divCalculations = CreateCalculationArray();
            var shapeCalculations = CreateCalculationArray();
            var fractionCalculations = CreateCalculationArray();
            var measureCalculations = CreateCalculationArray();

            DoMath(addCalculations, subCalculations, mulCalculations, divCalculations, shapeCalculations, fractionCalculations, measureCalculations);

            Console.WriteLine($"Math Done:{DateTime.Now:hh:mm:ss:f}");

            DebugConsole(addCalculations, subCalculations, mulCalculations, divCalculations, shapeCalculations, fractionCalculations, measureCalculations);

            OutputAndFileExchange(addCalculations, subCalculations, mulCalculations, divCalculations, fractionCalculations, measureCalculations);
        }

        private static void InitializeConsole()
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("Hello my muse");
            Console.ResetColor();
            Console.WriteLine(" ");
            Console.WriteLine(" ");

            Console.WriteLine($"Initilization: {DateTime.Now:hh:mm:ss:f}");
            Console.WriteLine(" ");
        }

        private static (Operation, int, int, int, string)[] CreateCalculationArray()
        {
            return new (Operation, int, int, int, string)[workSheets];
        }

        private static void OutputAndFileExchange(
            (Operation Operation, int Op1, int Op2, int Product, string Note)[] addCalculations,
            (Operation Operation, int Op1, int Op2, int Product, string Note)[] subCalculations,
            (Operation Operation, int Op1, int Op2, int Product, string Note)[] mulCalculations,
            (Operation Operation, int Op1, int Op2, int Product, string Note)[] divCalculations,
            (Operation Operation, int Op1, int Op2, int Product, string Note)[] fractionCalculations,
            (Operation Operation, int Op1, int Op2, int Product, string Note)[] measureCalculations)
        {
            targetPath = targetPath.Replace("@TimeStamp", DateTime.Now.ToString("yyyy_MM_dd.sf"));

            Console.WriteLine($"Prep Done:{DateTime.Now:hh:mm:ss:f}");

            for (int i = 0; i < workSheets; i++)
            {
                string dayName = Days[i];

                var op1AddStr = addCalculations[i].Op1.ToString("D4");
                var op2AddStr = addCalculations[i].Op2.ToString("D4");

                var op1SubStr = subCalculations[i].Op1.ToString("D4");
                var op2SubStr = subCalculations[i].Op2.ToString("D4");

                var problemMap = new Dictionary<string, string>
                {
                    ["^Add1.4"] = op1AddStr[0].ToString(),
                    ["^Add1.3"] = op1AddStr[1].ToString(),
                    ["^Add1.2"] = op1AddStr[2].ToString(),
                    ["^Add1.1"] = op1AddStr[3].ToString(),

                    ["^Add2.4"] = op2AddStr[0].ToString(),
                    ["^Add2.3"] = op2AddStr[1].ToString(),
                    ["^Add2.2"] = op2AddStr[2].ToString(),
                    ["^Add2.1"] = op2AddStr[3].ToString(),

                    ["^Sub1.4"] = op1SubStr[0].ToString(),
                    ["^Sub1.3"] = op1SubStr[1].ToString(),
                    ["^Sub1.2"] = op1SubStr[2].ToString(),
                    ["^Sub1.1"] = op1SubStr[3].ToString(),

                    ["^Sub2.4"] = op2SubStr[0].ToString(),
                    ["^Sub2.3"] = op2SubStr[1].ToString(),
                    ["^Sub2.2"] = op2SubStr[2].ToString(),
                    ["^Sub2.1"] = op2SubStr[3].ToString(),

                    ["^Multi1"] = mulCalculations[i].Op1.ToString(),
                    ["^Multi2"] = mulCalculations[i].Op2.ToString(),

                    ["^Div1"] = divCalculations[i].Op1.ToString(),
                    ["^Div2"] = divCalculations[i].Op2.ToString(),

                    ["^Fraction1.1"] = fractionCalculations[i].Op1.ToString(),
                    ["^Fraction1.2"] = fractionCalculations[i].Op2.ToString()
                };

                var tmpsourcePath = sourcePath.Replace("@DocName", $"{dayName}.docx");
                var tmptargetPath = targetPath.Replace("@DocName", $"{dayName}.docx");

                FillDocxFile(tmpsourcePath, tmptargetPath, problemMap);

                Console.WriteLine($"{dayName} Done:{DateTime.Now:hh:mm:ss:f}");
            }
        }


        #region Debug
        private static void DebugConsole(
            (Operation Operation, int Op1, int Op2, int Product, string Note)[] addCalculations,
            (Operation Operation, int Op1, int Op2, int Product, string Note)[] subCalculations,
            (Operation Operation, int Op1, int Op2, int Product, string Note)[] mulCalculations,
            (Operation Operation, int Op1, int Op2, int Product, string Note)[] divCalculations,
            (Operation Operation, int Op1, int Op2, int Product, string Note)[] shapeCalculations,
            (Operation Operation, int Op1, int Op2, int Product, string Note)[] fractionCalculations,
            (Operation Operation, int Op1, int Op2, int Product, string Note)[] measureCalculations)
        {
            for (int i = 0; i < workSheets; i++)
            {
                string dayName = Days[i];
                Console.ResetColor();
                Console.WriteLine($"Worksheet: {dayName}");

                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine($"{addCalculations[i].Operation} || {addCalculations[i].Op1:D4} + {addCalculations[i].Op2:D4}  = {addCalculations[i].Product:D4}");

                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine($"{subCalculations[i].Operation} || {subCalculations[i].Op1:D4} - {subCalculations[i].Op2:D4}  = {subCalculations[i].Product:D4}");

                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine($"{mulCalculations[i].Operation}  || {mulCalculations[i].Op1:D4} * {mulCalculations[i].Op2:D4}  = {mulCalculations[i].Product:D4}");

                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine($"{divCalculations[i].Operation}  || {divCalculations[i].Op1:D4} / {divCalculations[i].Op2:D4}  = {divCalculations[i].Product:D4}");

                Console.ForegroundColor = ConsoleColor.Magenta;
                Console.WriteLine($"{shapeCalculations[i].Operation}  || Shape with L {shapeCalculations[i].Op1:D2} and W {shapeCalculations[i].Op2:D2} | {shapeCalculations[i].Note}");

                Console.ForegroundColor = ConsoleColor.DarkBlue;
                Console.WriteLine($"{fractionCalculations[i].Operation}  || Fraction {shapeCalculations[i].Op1}/{shapeCalculations[i].Op2:D2}");

                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.WriteLine($"{shapeCalculations[i].Operation}  || Shape with L {shapeCalculations[i].Op1:D2} and W {shapeCalculations[i].Op2:D2} | {shapeCalculations[i].Note}");

                Console.WriteLine();
            }
        }
        #endregion

        #region Math
        private static void DoMath(
            (Operation Operation, int Op1, int Op2, int Product, string Note)[] addCalculations,
            (Operation Operation, int Op1, int Op2, int Product, string Note)[] subCalculations,
            (Operation Operation, int Op1, int Op2, int Product, string Note)[] mulCalculations,
            (Operation Operation, int Op1, int Op2, int Product, string Note)[] divCalculations,
            (Operation Operation, int Op1, int Op2, int Product, string Note)[] shapeCalculations,
            (Operation Operation, int Op1, int Op2, int Product, string Note)[] fractionCalculations,
            (Operation Operation, int Op1, int Op2, int Product, string Note)[] measureCalculations)
        {
            for (int i = 0; i < workSheets; i++)
            {
                addCalculations[i] = CreateAddProblem();
                subCalculations[i] = CreateSubProblem();
                mulCalculations[i] = CreateMulProblem();
                divCalculations[i] = CreateDivProblem();
                shapeCalculations[i] = CreateShapeProblem();
                fractionCalculations[i] = CreateFractionProblem();
                measureCalculations[i] = CreateMeasureProblem();
            }
        }

        private static (Operation Op, int Op1, int Op2, int Product, string Note) CreateAddProblem()
        {
            var addParm1 = GetRandomValue(randomCapped4500);
            var addParm2 = GetRandomValue(randomCapped4500);
            var addAnswer = addParm1 + addParm2;

            return (Operation.Addition, addParm1, addParm2, addAnswer, string.Empty);
        }

        private static (Operation Op, int Op1, int Op2, int Product, string Note) CreateSubProblem()
        {
            var subParm1 = GetRandomValue(randomCapped4500);
            var subParm2 = GetRandomValue(randomCapped4500);

            NudgeSubtraction();

            var subAnswer = subParm1 - subParm2;

            return (Operation.Subtract, subParm1, subParm2, subAnswer, string.Empty);

            void NudgeSubtraction()
            {
                if (subParm1 < subParm2)
                {
                    var temp = subParm1;
                    subParm1 = subParm2;
                    subParm2 = temp;
                }
            }
        }

        private static (Operation Op, int Op1, int Op2, int Product, string Note) CreateShapeProblem()
        {
            var shapeLength = GetRandomValue(randomCapped12);
            var shapeWidth = GetRandomValue(randomCapped12);
            var shapeArea = shapeLength * shapeWidth;
            var shapePerimeter = shapeLength * 2 + shapeWidth * 2;

            return (Operation.Shape, shapeLength, shapeWidth, 0, $"Area: {shapeArea} | Perimeter: {shapePerimeter}");
        }

        private static (Operation Op, int Op1, int Op2, int Product, string Note) CreateMulProblem()
        {
            var mulParm1 = GetRandomValue(randomCapped12);
            var mulParm2 = GetRandomValue(randomCapped12);
            var mulAnswer = mulParm1 * mulParm2;

            return (Operation.Multiply, mulParm1, mulParm2, mulAnswer, string.Empty);
        }

        private static (Operation Op, int Op1, int Op2, int Product, string Note) CreateMeasureProblem()
        {
            var measureParm1 = GetRandomValue(measurementItems);
            var measureParm2 = GetRandomValue(measurementItems);

            return (Operation.Multiply, 0, 0, 0, $"Measure items: {measureParm1}, {measureParm2}");
        }

        private static (Operation Op, int Op1, int Op2, int Product, string Note) CreateDivProblem()
        {
            var divParm2 = GetRandomValue(randomCapped12);
            var quotient = GetRandomValue(randomCapped12);
            var divParm1 = divParm2 * quotient;
            var divAnswer = quotient;

            return (Operation.Divide, divParm1, divParm2, divAnswer, string.Empty);
        }

        private static (Operation Op, int Op1, int Op2, int Product, string Note) CreateFractionProblem()
        {
            var addParm1 = GetRandomValue(randomDenominator);
            var addParm2 = GetRandomValue(randomDenominator);

            NudgeFraction();

            while (addParm2 == addParm1 )
            {
               addParm2 = Math.Abs(GetRandomValue(randomDenominator) - addParm2);
            }

            return (Operation.Fraction, addParm1, addParm2, 0, string.Empty);

            void NudgeFraction()
            {
                if (addParm1 > addParm2)
                {
                    var temp = addParm1;
                    addParm1 = addParm2;
                    addParm2 = temp;
                }
            }
        }

        public static void FillDocxFile(string sourcePath, string targetPath, Dictionary<string, string> pPmap)
        {

            if (!File.Exists(sourcePath))
                throw new FileNotFoundException("Source file not found.", sourcePath);

            // Ensure target directory exists
            var targetDir = Path.GetDirectoryName(targetPath);
            if (!string.IsNullOrEmpty(targetDir) && !Directory.Exists(targetDir))
                Directory.CreateDirectory(targetDir);

            File.Copy(sourcePath, targetPath, true);

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(targetPath, true))
            {
                var texts = wordDoc.MainDocumentPart.Document.Body.Descendants<Text>();
                foreach (var text in texts)
                {
                    foreach (var key in pPmap.Keys)
                    {
                        if (text.Text.Contains(key))
                        {
                            text.Text = text.Text.Replace(key, pPmap[key]);
                        }
                    }
                }
                wordDoc.MainDocumentPart.Document.Save();
            }
        }
        private static int GetRandomValue(int[] pArr)
        {
            return pArr[rnd.Next(1000) % pArr.Length];
        }

        private static string GetRandomValue(List<string> pArr)
        {
            return pArr[rnd.Next(pArr.Count)];
        }
        #endregion
    }
}