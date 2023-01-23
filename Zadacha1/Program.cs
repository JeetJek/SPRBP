// See https://aka.ms/new-console-template for more information
Console.WriteLine("Hello, World!");
if (args.Length < 1)
{
    Console.WriteLine("Не заданы пути!\nЗапустите заново указав пути для шаблона и результата.");
    return;
}
Zadacha1.WordFormatter wordFormatter = new Zadacha1.WordFormatter(args[0], args[1]);
wordFormatter.Make();