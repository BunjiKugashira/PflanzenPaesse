namespace PflanzenPaesse.Validators
{
    using System;
    using System.IO;
    using System.Linq;

    using PflanzenPaesse.Repositories.ExcelRepository;

    using PowerArgs;

    public class FileInputValidator : ArgValidator
    {
        public override void Validate(string name, ref string arg)
        {
            arg = FixPath(arg);

            if (!File.Exists(arg))
            {
                throw new FileNotFoundException(null, arg, null);
            }

            if (ExcelRepository.AllowedFileEndings.Contains(arg[arg.LastIndexOf('.')..]))
            {
                throw new ArgumentException("File format must be one of the following: " + string.Join(", ", ExcelRepository.AllowedFileEndings), name, null);
            }
        }

        public static string FixPath(string path)
        {
            if (path.StartsWith('"') && path.EndsWith('"') && !path.EndsWith("\\\"") && path.Length >= 2)
            {
                path = path[1..^1];
            }

            return path;
        }
    }
}
