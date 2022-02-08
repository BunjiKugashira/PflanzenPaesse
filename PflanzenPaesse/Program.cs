namespace PflanzenPaesse
{
    using System;
    using System.Threading.Tasks;

    using PowerArgs;

    public static class Program
    {
        private static async Task Main(string[] args)
        {
            try
            {
                await Args.InvokeMainAsync<ProgramArgs>(args);
                Console.WriteLine("Finished");
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(e);
                Console.ResetColor();
            }
            finally
            {
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
            }
        }
    }
}
