using System;


namespace CalculoCPF
{
    class Program
    {





     
        static Boolean checkCpf(string cpf){
            string tempCpf;
		   string digito;
		   int soma;
		   int resto;

            int[] mult1 = new int[9] { 10, 9, 8, 7, 6, 5, 4, 3, 2 };
		    int[] mult2 = new int[10] { 11, 10, 9, 8, 7, 6, 5, 4, 3, 2 };
            	   
		tempCpf = cpf.Substring(0, 9);
		soma = 0;
        for(int i=0; i<9; i++){
		    soma += int.Parse(tempCpf[i].ToString()) * mult1[i];
        }
		resto = soma % 11;
		if ( resto < 2 ){
             resto = 0;
        }else
		   resto = 11 - resto;
		digito = resto.ToString();
		tempCpf = tempCpf + digito;
		soma = 0;
		for(int i=0; i<10; i++)
		    soma += int.Parse(tempCpf[i].ToString()) * mult2[i];
		resto = soma % 11;
		if (resto < 2)
		   resto = 0;
		else
		   resto = 11 - resto;
		digito = digito + resto.ToString();
		return cpf.EndsWith(digito);
        }




        static String preencher(){
            int aux=0;
            string cpf="";
             while(aux==0){

           
            if(cpf.Length>11||cpf.Length<11){
                Console.Write("Digite cpf sem traços ou pontos: ");
                cpf=Console.ReadLine();
            }else{
                aux=1;
            }
            
          
            
            }
              return cpf;
        }

        /*
        1 – Cadastro de clientes
2 – Cadastro de produtos
3 – Realizar venda
4 – Extrato cliente
9 – Sair

 */
        static void menu(){
            int aux=0;
            while(aux!=9){
                Console.WriteLine("1:CADASTRO DE CLIENTES");
                Console.WriteLine("2:CADASTRO DE PRODUTOS");
                Console.WriteLine("3:REALIZAR VENDA");
                Console.WriteLine("4:EXTRATO CLIENTE");
                Console.WriteLine("9:SAIR");
                aux=Int32.Parse(Console.ReadLine());

            }
            Console.WriteLine("-----------------------Operaçã finalizada-----------------------");
        }







        static void Main(string[] args)
        {
            string cpf="";
            int[] numeros=new int[11];
            
            Console.Write("Digite o cpf: ");
            cpf=Console.ReadLine();
            cpf=preencher();

            checkCpf(cpf);
            Console.WriteLine(checkCpf(cpf));

            


      
           

        }


    }
}
