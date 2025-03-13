using DesafioDosMedicos.Modelos;
using NPOI.SS.UserModel;
using Org.BouncyCastle.Crypto.Engines;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.Math.EC.Rfc7748;
using NPOI.OpenXmlFormats;
using ICSharpCode.SharpZipLib.Zip;
using System.Text.RegularExpressions;
using NPOI.SS.Formula.Functions;
using NPOI.HSSF.Record;
using MathNet.Numerics;

internal class Program
{

    private static string caminhoPlanilha = Path.Combine(Environment.CurrentDirectory, "DesafioMedicos.xlsx");
    private static List<Consulta> listaDeConsulta = [];
    static void Main(string[] args)
    {
        ImportarDadosPlanilha();
        // ExibirDadosDaPlanilha();

        // ExercicioUm();
        // ExercicioDois();
        // ExercicioTres();
        // ExercicioQuatro();
        // ExercicioCinco();
        Desafio1();
    }
    private static void ImportarDadosPlanilha()
    {
        try
        {
            IWorkbook pastaTrabalho = WorkbookFactory.Create(caminhoPlanilha);
            ISheet planilha = pastaTrabalho.GetSheetAt(0);
            for (int i = 1; i < planilha.PhysicalNumberOfRows; i++)
            {
                IRow linha = planilha.GetRow(i);
                DateTime dataConsulta = DateTime.Parse(linha.GetCell(0).StringCellValue);
                string horaDaConsulta = linha.GetCell(1).StringCellValue;
                string nomePaciente = linha.GetCell(2).StringCellValue;
                string? numeroTelefone = linha.GetCell(3)?.StringCellValue;
                long cpf = Convert.ToInt64(Regex.Replace(linha.GetCell(4).StringCellValue, @"\D", ""));
                string rua = linha.GetCell(5).StringCellValue;
                string cidade = linha.GetCell(6).StringCellValue;
                string estado = linha.GetCell(7).StringCellValue;
                string especialidade = linha.GetCell(8).StringCellValue;
                string nomeMedico = linha.GetCell(9).StringCellValue;
                bool particular = linha.GetCell(10).StringCellValue == "Sim" ? true : false;
                long numeroDaCarteirinha = (long)linha.GetCell(11).NumericCellValue;
                double valorDaConsulta = linha.GetCell(12).NumericCellValue;

                Consulta consulta = new(dataConsulta, horaDaConsulta, nomePaciente, numeroTelefone, cpf, rua, cidade, estado, especialidade, nomeMedico, particular, numeroDaCarteirinha, valorDaConsulta);
                listaDeConsulta.Add(consulta);

            }

        }
        catch (Exception erro)
        {

            System.Console.WriteLine(erro.Message);
        }

    }

    static void ExibirDadosDaPlanilha()
    {
        // var nomePacientes = listaDeConsulta.Select(x => x.NomePaciente).ToList();
        foreach (var nomes in listaDeConsulta)
        {
            Console.WriteLine(nomes);
        }
    }

    //1 – Liste ao total quantos pacientes temos para atender do dia 27/03 até dia 31/03. Sem repetições.
    private static void ExercicioUm()
    {
        var totalDePacientes = listaDeConsulta.Where(x => (x.DataConsulta.Day >= 27 && x.DataConsulta.Month == 3)
        && (x.DataConsulta.Day <= 31 && x.DataConsulta.Month == 3)).DistinctBy(g => g.NomePaciente).Count();
        System.Console.WriteLine(totalDePacientes);
    }
    //2 – Liste ao total quantos médicos temos trabalhando em nosso consultório. Conte a quantidade de médicos sem repetições. 
    private static void ExercicioDois()
    {
        var quantidadeTotalMedicos = listaDeConsulta.Select(x => x.NomeMedico).Distinct().Count();

        System.Console.WriteLine(quantidadeTotalMedicos);
    }
    //3 – Liste o nome dos médicos e suas especialidades.
    private static void ExercicioTres()
    {
        var nomeMedicoEspecialidade = listaDeConsulta.Select(x => new
        {
            NomeMedico = x.NomeMedico,
            EspecialidadeMedico = x.Especialidade

        });

        foreach (var medico in nomeMedicoEspecialidade)
        {
            System.Console.WriteLine($"Medico: {medico.NomeMedico} | Especialidade: {medico.EspecialidadeMedico}");
        }

    }

    //4 – Liste o total em valor de consulta que receberemos. Some o valor de todas as consultas.
    // Depois liste o valor por especialidade.
    private static void ExercicioQuatro()
    {
        var valorConsultEspecialidade = listaDeConsulta.GroupBy(x => x.Especialidade).Select(g => new
        {
            especialidade = g.Key,
            totalEspecialidade = g.Sum(x => x.ValorDaConsulta)
        });
        System.Console.WriteLine($"\nValor total: {valorConsultEspecialidade.Sum(x => x.totalEspecialidade):c}\n");
        foreach (var item in valorConsultEspecialidade)
        {
            System.Console.WriteLine($"{item.especialidade} | {item.totalEspecialidade:c}");
        }
    }

    //5 – Para o dia 30/03. Quantas consultas vão ser realizadas? Quantas são Particular?
    //  Liste para esse dia os horários de consulta de cada médico e suas especialidades.

    private static void ExercicioCinco()
    {
        var consultasRealziadas = listaDeConsulta.Where(x => (x.DataConsulta.Day == 30 && x.DataConsulta.Month == 3)).
        GroupBy(a => new
        {
            nomeMedico = a.NomeMedico,
            especialidadeMedico = a.Especialidade,
        }).Select(g => new
        {
            Medico = g.Key.nomeMedico,
            Especialidade = g.Key.especialidadeMedico,
            horario = g.Select(c => c.HoraDaConsulta),
            consultasParticulares = g.Count(b => b.Particular == true)
        });

        var totalConsutas = consultasRealziadas.Sum(x => x.horario.Count());
        var totalConsultasParticulares = consultasRealziadas.Sum(g => g.consultasParticulares);

        System.Console.WriteLine($"Total de consultas no dia 30/03: {totalConsutas}\nTotal de consultas particulares no dia 30/03: {totalConsultasParticulares}");

        foreach (var item in consultasRealziadas)
        {
            System.Console.WriteLine($"Medico: {item.Medico} | Especialidade: {item.Especialidade}");
            foreach (var horario in item.horario)
            {
                System.Console.WriteLine($"Horario {horario}");
            }
        }
    }

    //1 – Verifique se algum paciente tem alguma consulta marcada no mesmo horário.
    // Tem? Aponte quais, pois precisaremos ligar para o paciente.
    // Não tem telefone? Procure se há alguém que more na mesma Rua, Cidade e Estado que o paciente para tentarmos entrar em contato.
    private static void Desafio1()
    {
        var consultasNaMesmaDataHora = listaDeConsulta.GroupBy(x => new
        {
            horario = x.HoraDaConsulta,
            data = x.DataConsulta,
            paciente = x.Cpf
        })
        .Where(g => g.Select(a => a.Cpf).Count() > 1)
       .ToList();

        foreach (var item in consultasNaMesmaDataHora)
        {
            System.Console.WriteLine(item);
            System.Console.WriteLine("...");
        }
    }
}