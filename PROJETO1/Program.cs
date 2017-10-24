using System;
using System.IO;
using NetOffice.ExcelApi;

namespace PROJETO1
{
    class Program
    {
        public static bool checkCnpj(string cnpj)
        {
            int[] multiplicador1 = new int[12] { 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2 };
            int[] multiplicador2 = new int[13] { 6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2 };
            int soma;
            int resto;
            string digito;
            string tempCnpj;
            if (cnpj.Length != 14)
                return false;
            tempCnpj = cnpj.Substring(0, 12);
            soma = 0;
            for (int i = 0; i < 12; i++)
                soma += int.Parse(tempCnpj[i].ToString()) * multiplicador1[i];
            resto = (soma % 11);
            if (resto < 2)
                resto = 0;
            else
                resto = 11 - resto;
            digito = resto.ToString();
            tempCnpj = tempCnpj + digito;
            soma = 0;
            for (int i = 0; i < 13; i++)
                soma += int.Parse(tempCnpj[i].ToString()) * multiplicador2[i];
            resto = (soma % 11);
            if (resto < 2)
                resto = 0;
            else
                resto = 11 - resto;
            digito = digito + resto.ToString();
            return cnpj.EndsWith(digito);
        }




        static Boolean checkCpf(string cpf)
        {
            string tempCpf;
            string digito;
            int soma;
            int resto;

            int[] mult1 = new int[9] { 10, 9, 8, 7, 6, 5, 4, 3, 2 };
            int[] mult2 = new int[10] { 11, 10, 9, 8, 7, 6, 5, 4, 3, 2 };

            tempCpf = cpf.Substring(0, 9);
            soma = 0;
            for (int i = 0; i < 9; i++)
            {
                soma += int.Parse(tempCpf[i].ToString()) * mult1[i];
            }
            resto = soma % 11;
            if (resto < 2)
            {
                resto = 0;
            }
            else
                resto = 11 - resto;
            digito = resto.ToString();
            tempCpf = tempCpf + digito;
            soma = 0;
            for (int i = 0; i < 10; i++)
                soma += int.Parse(tempCpf[i].ToString()) * mult2[i];
            resto = soma % 11;
            if (resto < 2)
                resto = 0;
            else
                resto = 11 - resto;
            digito = digito + resto.ToString();
            return cpf.EndsWith(digito);
        }




        static String preencher()
        {
            int aux = 0;
            string cpf = "";
            while (aux == 0)
            {


                if (cpf.Length > 11 || cpf.Length < 11)
                {
                    Console.Write("Digite cpf sem traços ou pontos: ");
                    cpf = Console.ReadLine();
                }
                else
                {
                    aux = 1;
                }



            }
            return cpf;
        }

        static String preencherCnpj()
        {
            int aux = 0;
            string cpf = "";
            while (aux == 0)
            {


                if (cpf.Length > 14 || cpf.Length < 14)
                {
                    Console.Write("Digite cnpj sem traços ou pontos: ");
                    cpf = Console.ReadLine();
                }
                else
                {
                    aux = 1;
                }



            }
            return cpf;
        }

        static void venda()
        {
            Application excel = new Application();

            FileInfo cadCliente = new
            FileInfo(@"C:\Users\Admnin2\Desktop\projetos\PROJETO1\cadCliente.xlsx");
            FileInfo cadProduto = new
            FileInfo(@"C:\Users\Admnin2\Desktop\projetos\PROJETO1\cadProduto.xlsx");
            FileInfo cadVenda = new
            FileInfo(@"C:\Users\Admnin2\Desktop\projetos\PROJETO1\cadVenda.xlsx");
            string[,] clientes = new string[500, 4];
            string[,] produtos = new string[500, 4];
            excel.Workbooks.Open(@"C:\Users\Admnin2\Desktop\projetos\PROJETO1\cadCliente.xlsx");
            excel.Range("A1").Select();

            
             Console.WriteLine("Pessoa Física ou Jurídica?");
        string cpf = Console.ReadLine().ToLower();
        if (cpf  == "fisica")
        {
            cpf  = preencher();
            checkCpf(cpf );
        }
        else if (cpf  == "juridica")
        {
            cpf  = preencherCnpj();

            checkCnpj(cpf );

        }

            for (int lin = 0; lin < 500; lin++)
            {
                for (int col = 0; col < 4; col++)
                {
                    if (excel.Range("A" + lin).Value == null)
                    {
                        col += 4;
                        lin += 500;

                    }
                    else
                    {clientes[lin, col] = excel.Cells[lin + 1, col + 1].Value.ToString();

                        

                    }

                }
                if (clientes[lin, 2] != cpf)
                {
                    menu();

                }
            }


            
            excel.Quit();
            excel.Workbooks.Open(@"C:\Users\Admnin2\Desktop\projetos\PROJETO1\cadProduto.xlsx");
            excel.Range("A1").Select();
            for (int lin = 0; lin < 500; lin++)
            {
                for (int col = 0; col < 4; col++)
                {
                    if (excel.Range("A" + lin).Value != null)
                    {
                        produtos[lin, col] = excel.Cells[lin + 1, col + 1].Value.ToString();



                    }
                    else
                    {
                        col += 4;
                        lin += 500;

                    }

                }
            }
            excel.Quit();

            for (int lin = 0; lin < 500; lin++)
            {
                for (int col = 0; col < 4; col++)
                {
                    Console.Write(produtos[lin, col] + "\t");

                }
                Console.WriteLine();
                if (produtos[lin, 0] == null)
                {
                    lin += 500;
                }
            }


            string prod;

            Console.Write("Digite o numero do produto: ");
            prod = Console.ReadLine();
             if (cadCliente.Exists)
        {

            excel.Workbooks.Open(@"C:\Users\Admnin2\Desktop\projetos\PROJETO1\cadCliente.xlsx");
            excel.Range("A1").Select();
            for (int i = 1; i < 1000; i++)
            {
                if (excel.Range("A" + i).Value == null)
                {
                    excel.Range("A" + i).Value = cpf;
                    excel.Range("B" + i).Value = prod;
                    excel.Range("C" + i).Value = DateTime.Now.ToShortDateString();
                    excel.ActiveWorkbook.Save();
                    Console.Clear();
                    Console.WriteLine("Cadastro realizado com sucesso");
                    Console.WriteLine("Aperte uma tecla para continuar");
                    Console.ReadLine();
                    Console.Clear();
                    break;
                }
            }
        }
        else
        {
            excel.Workbooks.Add();
            excel.Range("A1").Select();
            excel.Range("A1").Value = "cpf/cnpj";
            excel.Range("B1").Value = "Codigo de produto";
            excel.Range("C1").Value = "Data de Cadastro";

            excel.Range("A2").Value = cpf;
            excel.Range("B2").Value = prod;

            excel.Range("C2").Value = DateTime.Now.ToShortDateString();
            Console.WriteLine("Cadastro realizado com sucesso");
            Console.WriteLine("Aperte uma tecla para continuar");
            Console.ReadLine();
            Console.Clear();


            excel.ActiveWorkbook.SaveAs(@"C:\Users\Admnin2\Desktop\projetos\PROJETO1\cadVenda.xlsx");
        }
        excel.Quit();




        }

    








    


    /*
    1 – Cadastro de clientes
2 – Cadastro de produtos
3 – Realizar venda
4 – Extrato cliente
9 – Sair

*/
    static void menu()
    {
        int aux = 0;
        while (aux != 9)
        {


            Console.WriteLine("1:CADASTRO DE CLIENTES");
            Console.WriteLine("2:CADASTRO DE PRODUTOS");
            Console.WriteLine("3:REALIZAR VENDA");
            Console.WriteLine("4:EXTRATO CLIENTE");
            Console.WriteLine("9:SAIR");
            aux = Int32.Parse(Console.ReadLine());



            switch (aux)
            {
                case 1: CadastroCli(); break;

                case 2: CadastroProd(); break;
                case 3: venda(); break;



                default: Console.WriteLine("OPÇÂO INVALIDA"); break;





            }

        }
        Console.WriteLine("-----------------------Operaçã finalizada-----------------------");
    }



    static void CadastroProd()
    {
        Application excel = new Application();
        FileInfo cadProduto = new
       FileInfo(@"C:\Users\Admnin2\Desktop\projetos\PROJETO1\cadProduto.xlsx");


        Console.WriteLine("Digite o codigo do produto:");
        string codigo = Console.ReadLine();
        Console.WriteLine("Digite o nome do produto");
        string nome = Console.ReadLine();
        Console.WriteLine("Digite a descrição do produto:");
        string descri = Console.ReadLine();
        Console.WriteLine("Digite o preço do produto");
        string preco = Console.ReadLine();
        if (cadProduto.Exists)
        {

            excel.Workbooks.Open(@"C:\Users\Admnin2\Desktop\projetos\PROJETO1\cadProduto.xlsx");
            excel.Range("A1").Select();
            for (int i = 1; i < 1000; i++)
            {
                if (excel.Range("A" + i).Value == null)
                {
                    excel.Range("A" + i).Value = codigo;
                    excel.Range("B" + i).Value = nome;
                    excel.Range("C" + i).Value = descri;
                    excel.Range("D" + i).Value = preco;
                    excel.ActiveWorkbook.Save();
                    Console.Clear();
                    Console.WriteLine("Cadastro realizado com sucesso");
                    Console.WriteLine("Aperte uma tecla para continuar");
                    Console.ReadLine();
                    Console.Clear();
                    break;
                }
            }
        }
        else
        {
            excel.Workbooks.Add();
            excel.Range("A1").Select();
            excel.Range("A1").Value = "Codigo";
            excel.Range("B1").Value = "Nome";
            excel.Range("C1").Value = "Descrição";
            excel.Range("D1").Value = "Preço";

            excel.Range("A2").Value = codigo;
            excel.Range("B2").Value = nome;

            excel.Range("C2").Value = descri;
            excel.Range("D2").Value = preco;
            Console.WriteLine("Cadastro realizado com sucesso");
            Console.WriteLine("Aperte uma tecla para continuar");
            Console.ReadLine();
            Console.Clear();


            excel.ActiveWorkbook.SaveAs(@"C:\Users\Admnin2\Desktop\projetos\PROJETO1\cadProduto.xlsx");
        }
        excel.Quit();



    }




    static void CadastroCli()
    {
        Application excel = new Application();

        FileInfo cadCliente = new
        FileInfo(@"C:\Users\Admnin2\Desktop\projetos\PROJETO1\cadCliente.xlsx");


        Console.WriteLine("Digite o nome do cliente:");
        string nome = Console.ReadLine();
        Console.WriteLine("Digite o email do cliente");
        string email = Console.ReadLine();
        Console.WriteLine("Pessoa Física ou Jurídica?");
        string pessoa = Console.ReadLine().ToLower();
        if (pessoa == "fisica")
        {
            pessoa = preencher();
            checkCpf(pessoa);
        }
        else if (pessoa == "juridica")
        {
            pessoa = preencherCnpj();

            checkCnpj(pessoa);

        }

        if (cadCliente.Exists)
        {

            excel.Workbooks.Open(@"C:\Users\Admnin2\Desktop\projetos\PROJETO1\cadCliente.xlsx");
            excel.Range("A1").Select();
            for (int i = 1; i < 1000; i++)
            {
                if (excel.Range("A" + i).Value == null)
                {
                    excel.Range("A" + i).Value = nome;
                    excel.Range("B" + i).Value = email;
                    excel.Range("C" + i).Value = pessoa;
                    excel.Range("D" + i).Value = DateTime.Now.ToShortDateString();
                    excel.ActiveWorkbook.Save();
                    Console.Clear();
                    Console.WriteLine("Cadastro realizado com sucesso");
                    Console.WriteLine("Aperte uma tecla para continuar");
                    Console.ReadLine();
                    Console.Clear();
                    break;
                }
            }
        }
        else
        {
            excel.Workbooks.Add();
            excel.Range("A1").Select();
            excel.Range("A1").Value = "Nome";
            excel.Range("B1").Value = "E-mail";
            excel.Range("C1").Value = "CPF/CNPJ";
            excel.Range("D1").Value = "Data de Cadastro";

            excel.Range("A2").Value = nome;
            excel.Range("B2").Value = email;

            excel.Range("C2").Value = pessoa;
            excel.Range("D2").Value = DateTime.Now.ToShortDateString();
            Console.WriteLine("Cadastro realizado com sucesso");
            Console.WriteLine("Aperte uma tecla para continuar");
            Console.ReadLine();
            Console.Clear();


            excel.ActiveWorkbook.SaveAs(@"C:\Users\Admnin2\Desktop\projetos\PROJETO1\cadCliente.xlsx");
        }
        excel.Quit();






    }







    static void Main(string[] args)
    {

        Application excel = new Application();

        FileInfo cadCliente = new
        FileInfo(@"C:\Users\Admnin2\Desktop\projetos\PROJETO1\cadCliente.xlsx");
        FileInfo cadProduto = new
        FileInfo(@"C:\Users\Admnin2\Desktop\projetos\PROJETO1\cadProduto.xlsx");
        FileInfo cadVenda = new
        FileInfo(@"C:\Users\Admnin2\Desktop\projetos\PROJETO1\cadVenda.xlsx");
        menu();
        excel.Quit();
    }
    }
}







    




