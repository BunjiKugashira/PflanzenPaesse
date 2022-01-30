namespace PflanzenPaesse.Validators
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using PflanzenPaesse.Repositories.ExcelRepository;

    using PowerArgs;

    public class NumberRangeValidator : ArgValidator
    {
        public override void Validate(string name, ref string arg)
        {
            Parse(arg, ExcelRepository.MinRowIndex, ExcelRepository.MinRowIndex, true);
        }

        public static IEnumerable<int> Parse(string arg, int rangeMin, int rangeMax, bool validateOnly = false)
        {
            var ranges = arg.Split(',');
            var rangeIndicators = ranges.Select(range => range.Split('-').Select(indicator => indicator.Trim()));

            foreach(var rangeIndicator in rangeIndicators)
            {
                if (rangeIndicator.Count() == 1)
                {
                    yield return int.Parse(rangeIndicator.Single());
                }

                if (rangeIndicator.Count() == 2)
                {
                    var start = rangeIndicator.First();
                    var end = rangeIndicator.Last();
                    var startIndex = string.IsNullOrEmpty(start) ? rangeMin : int.Parse(start);
                    var endIndex = string.IsNullOrEmpty(end) ? rangeMax : int.Parse(end);

                    if (startIndex < ExcelRepository.MinRowIndex)
                    {
                        throw new IndexOutOfRangeException($"Range start was set to {startIndex} but can't be lower than {ExcelRepository.MinRowIndex}.");
                    }

                    if (endIndex > ExcelRepository.MaxRowIndex)
                    {
                        throw new IndexOutOfRangeException($"Range end was set to {endIndex} but can't be higher than {ExcelRepository.MaxRowIndex}.");
                    }

                    for (var i = startIndex; i <= endIndex && !validateOnly; i++)
                    {
                        yield return i;
                    }
                }
            }
        }
    }
}
