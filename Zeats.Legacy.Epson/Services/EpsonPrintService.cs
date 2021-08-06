using System;
using System.Collections.Generic;
using Zeats.Legacy.PlainTextTable.Print.Enums;
using Zeats.Legacy.PlainTextTable.Print.Print;

namespace Zeats.Legacy.Epson.Services
{
    public class EpsonPrintService : IPrintService
    {
        private static readonly object Lock = new object();

        private static readonly Dictionary<string, bool> Printers = new Dictionary<string, bool>();

        public void Print(PrintCollection printCollection)
        {
            try
            {
                lock (Lock)
                {
                    TryOpenPort(printCollection);

                    foreach (var printItem in printCollection)
                    {
                        var tipoLetra = printItem.FontSize == FontSize.Small ? 1
                            : printItem.FontSize == FontSize.Normal ? 2
                            : printItem.FontSize == FontSize.Large ? 3
                            : 2;

                        var italico = printItem.Italic ? 1 : 0;
                        var sublinhado = printItem.Underline ? 1 : 0;
                        var expandido = printItem.FontSize == FontSize.Large ? 1 : 0;
                        var enfatizado = printItem.Bold ? 1 : 0;


                        int retorno;
                        if (printItem.FontType == FontType.Text)
                            retorno = InterfaceEpsonNF.FormataTX(printItem.Content, tipoLetra, italico, sublinhado, expandido, enfatizado);
                        else
                        {
                            InterfaceEpsonNF.ConfiguraCodigoBarras(100, 2, 2, 1, 20);

                            switch (printItem.BarCodeType)
                            {
                                case BarCodeType.Ean13:
                                    retorno = InterfaceEpsonNF.ImprimeCodigoBarrasEAN13(printItem.Content);
                                    break;

                                case BarCodeType.Code128:
                                    retorno = InterfaceEpsonNF.ImprimeCodigoBarrasCODE128(printItem.Content);
                                    break;

                                default:
                                    retorno = InterfaceEpsonNF.ImprimeCodigoBarrasEAN13(printItem.Content);
                                    break;
                            }
                        }

                        if (retorno == 0)
                            throw new Exception("Ocorreu um erro ao enviar um comando de impressão para a impressora");
                    }
                }
            }
            catch (Exception exception)
            {
                TryRelease(printCollection);
                Console.WriteLine(exception);
                throw;
            }
        }

        public void Cut(string portName, CutType cut = CutType.Full)
        {
            try
            {
                lock (Lock)
                {
                    TryOpenPort(portName);

                    InterfaceEpsonNF.AcionaGuilhotina(cut == CutType.Partial ? 0 : 1);
                }
            }
            catch (Exception exception)
            {
                TryRelease(portName);
                Console.WriteLine(exception);
                throw;
            }
        }

        private static void TryOpenPort(PrintCollection printCollection)
        {
            TryOpenPort(printCollection.Options.PortName);
        }

        private static void TryOpenPort(string portName)
        {
            if (Printers.ContainsKey(portName) && Printers[portName])
                return;

            var retorno = InterfaceEpsonNF.IniciaPorta(portName);
            if (retorno == 0)
                throw new Exception("Não foi possível estabelecer comunicação com a impressora");

            if (!Printers.ContainsKey(portName))
                Printers.Add(portName, true);

            Printers[portName] = true;
        }

        private static void TryRelease(PrintCollection printCollection)
        {
            TryRelease(printCollection.Options.PortName);
        }

        private static void TryRelease(string portName)
        {
            lock (Lock)
            {
                try
                {
                    if (Printers.ContainsKey(portName))
                        Printers[portName] = false;

                    InterfaceEpsonNF.FechaPorta();
                }
                catch (Exception exception)
                {
                    Console.WriteLine(exception);
                    throw;
                }
            }
        }
    }
}