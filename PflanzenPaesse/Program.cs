namespace PflanzenPaesse
{
    using System;

    using PowerArgs;

    public static class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Args.InvokeMain<ProgramArgs>(args);
                Console.WriteLine("Finished");
            }
            catch(Exception e)
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
