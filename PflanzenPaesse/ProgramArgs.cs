namespace PflanzenPaesse
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;

    using PflanzenPaesse.Repositories.ConstantsRepository;
    using PflanzenPaesse.Repositories.ExcelRepository;
    using PflanzenPaesse.Repositories.WordRepository;
    using PflanzenPaesse.Validators;

    using PowerArgs;

    [TabCompletion]
    public class ProgramArgs
    {
        [HelpHook, ArgShortcut("-?"), ArgDescription("Shows this help.")]
        public bool Help { get; set; }

        [ArgPosition(0), ArgRequired(PromptIfMissing = true), FileInputValidator, ArgShortcut("-i"), ArgDescription("The excel file used for input.")]
        public string InputFile { get; set; }

        [ArgPosition(1), ArgRequired(PromptIfMissing = true), ArgShortcut("-s"), ArgDescription("The sheet used for input.")]
        public string Sheet { get; set; }

        [ArgPosition(2), FileOutputValidator, ArgShortcut("-o"), ArgDescription("The word file used for output.")]
        public string OutputFile { get; set; }

        [ArgPosition(3), ArgDefaultValue("-"), ArgShortcut("-r"), ArgDescription("The rows to be processed.\nSeveral rows can be separated by ',' and ranges can be added by writing the first row (inclusive) and last row (inclusive) separated by '-'. A range can be left open on either side by not including a number."), ArgExample("1,2,5-7, 9-", "Process rows 1, 2, 5, 6, 7, 9 and all rows after that.")]
        public string Rows { get; set; }

        [ArgDefaultValue(3), ArgShortcut("-b"), ArgDescription("Insert a page-break after x passports.")]
        public uint MaxPassportsPerPage { get; set; }

        public async Task Main()
        {
            var outputFilePath = string.IsNullOrEmpty(this.OutputFile) ? FileOutputValidator.FixPath(this.InputFile) : this.OutputFile;

            Console.WriteLine($"InputFile: {this.InputFile}");
            Console.WriteLine($"OutputFile: {outputFilePath}");
            Console.WriteLine($"ProcessLines: {string.Join(", ", NumberRangeValidator.Parse(this.Rows, ExcelRepository.MinRowIndex, 10))}");

            var highestRowNumber = ExcelRepository.HighestUsedRowNumber(this.InputFile, this.Sheet);
            var rowsToProcess = NumberRangeValidator.Parse(this.Rows, 2, highestRowNumber);
            var data = ExcelRepository.Import(this.InputFile, this.Sheet).ToList();

            data = data.Where((row, id) => rowsToProcess.Contains(id + 2)).ToList();

            var constants = ConstantsRepository.Import();
            foreach (var datum in data)
            {
                foreach (var constant in constants)
                {
                    datum.TryAdd(constant.Key, constant.Value);
                }
            }

            await WordRepository.BuildPaesseAsync(@".\Templates\PflanzenpassTemplate.docx", data, outputFilePath, this.MaxPassportsPerPage);
        }
    }
}
