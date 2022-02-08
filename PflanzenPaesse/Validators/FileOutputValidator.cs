namespace PflanzenPaesse.Validators
{
    using System;
    using System.IO;
    using System.Linq;

    using PflanzenPaesse.Repositories.WordRepository;

    using PowerArgs;

    public class FileOutputValidator : ArgValidator
    {
        public override void Validate(string name, ref string arg)
        {
            arg = FixPath(arg);

            var file = new FileInfo(arg);

            if (!file.Directory.Exists)
            {
                throw new IOException($"Folder path \"{file.Directory}\" not found. Please create the folder or specify another folder for output.", null);
            }
        }

        public static string FixPath(string path, string ending = null)
        {
            if (path.StartsWith('"') && path.EndsWith('"') && !path.EndsWith("\\\"") && path.Length >= 2)
            {
                path = path[1..^1];
            }

            if (!path.Contains('.'))
            {
                path += '.';
            }

            if (!WordRepository.AllowedFileEndings.Contains(path.Substring(path.LastIndexOf('.'))))
            {
                path = path.Substring(0, path.LastIndexOf('.') + 1) + (ending ?? WordRepository.AllowedFileEndings[0]);
            }

            return Path.GetFullPath(path);
        }
    }
}
