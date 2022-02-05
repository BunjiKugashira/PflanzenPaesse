namespace PflanzenPaesse
{
    using System;
    using System.Linq;

    using PflanzenPaesse.Repositories.ExcelRepository;
    using PflanzenPaesse.Validators;

    using PowerArgs;

    [TabCompletion]
    public class ProgramArgs
    {        
        [HelpHook, ArgShortcut("-?"), ArgDescription("Shows this help.")]
        public bool Help { get; set; }

        [ArgPosition(0), ArgRequired(PromptIfMissing =true), FileInputValidator, ArgShortcut("-i"), ArgDescription("The excel file used for input.")]
        public string InputFile { get; set; }

        [ArgPosition(1), FileOutputValidator, ArgShortcut("-o"), ArgDescription("The word file used for output.")]
        public string OutputFile { get; set; }

        [ArgPosition(2), ArgDefaultValue("-"), ArgShortcut("-r"), ArgDescription("The rows to be processed.\nSeveral rows can be separated by ',' and ranges can be added by writing the first row (inclusive) and last row (inclusive) separated by '-'. A range can be left open on either side by not including a number."), ArgExample("1,2,5-7, 9-", "Process rows 1, 2, 5, 6, 7, 9 and all rows after that.")]
        public string Rows { get; set; }

        public void Main()
        {
            var outputFilePath = string.IsNullOrEmpty(OutputFile) ? FileOutputValidator.FixPath(InputFile) : OutputFile;

            Console.WriteLine($"InputFile: {InputFile}");
            Console.WriteLine($"OutputFile: {outputFilePath}");
            Console.WriteLine($"ProcessLines: {string.Join(", ", NumberRangeValidator.Parse(Rows, ExcelRepository.MinRowIndex, 10))}");

            var highestRowNumber = ExcelRepository.HighestUsedRowNumber(InputFile, "Januar");
            var rowsToProcess = NumberRangeValidator.Parse(Rows, 2, highestRowNumber);
            var data = ExcelRepository.Import(InputFile, "Januar").ToList();
        }
    }
}
