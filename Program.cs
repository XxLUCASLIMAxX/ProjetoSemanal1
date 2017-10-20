using System;


namespace CalculoCPF
{
    class Program
    {
        static Boolean checkCpf(int[]v){
            Boolean retorno= false;
            int[] n=new int[2];
            int soma=0 ;
            int mult=10;
            int resto=0;
            n[0]=v[9];
            
              
            n[1]=v[10];
             
            for(int i=0;i<9;i++){
                soma+=v[i]*mult;
               
                
                 mult--;
            }
            resto=soma%11;

            
           
         

            if(resto<2){
            
                resto=0;

                if(resto==n[0]){
                    
              
                    mult=11;
                    soma=0;
                    
                    for(int i =0;i<10;i++){
                        soma+=v[i]*mult;
                        mult--;
                    } 
                    resto=soma%11;
                    
                    if (resto<2){
                      
             
                        resto=0;
                        if(resto==n[1]){
                              Console.WriteLine("CPF VALIDO");
                            retorno=true;
                        }else{
                            Console.WriteLine("cpf invalido");
                        }
                    }else{
                        resto=11-resto;
                         
                        if(resto==n[1]){
                              Console.WriteLine("CPF VALIDO");
                            retorno=true;
                        }else{
                            Console.WriteLine("cpf invalido");
                        }
                    }

                }

            }else{
                resto=11-resto;
                
                
                if(resto==n[0]){
                     
                     mult=11;
                    soma=0;
                    
                    for(int i =0;i<10;i++){
                        soma+=v[i]*mult;
                        mult--;
                    } 
                    resto=soma%11;
                    if (resto<2){
                        resto=0;
                        if(resto==n[1]){
                            Console.WriteLine("CPF VALIDO");
                            retorno=true;
                        }else{
                            Console.WriteLine("cpf invalido");
                        }
                    }else{
                        resto=11-resto;
                        if(resto==n[1]){
                              Console.WriteLine("CPF VALIDO");
                            retorno=true;
                        }else{
                            Console.WriteLine("cpf invalido");
                        }
                   
                }
            }
       }

            return retorno;
        }
        static void preencher(string cpf,int[] numeros){
            int aux=0;
             while(aux==0){

           
            if(cpf.Length>11||cpf.Length<11){
                Console.Write("Digite cpf sem traços ou pontos: ");
                cpf=Console.ReadLine();
            }else{
                aux=1;
            }
            
            }
        }







        static void Main(string[] args)
        {
            string cpf="";
            int[] numeros=new int[11];
            
            Console.Write("Digite o cpf: ");
            cpf=Console.ReadLine();

            preencher(cpf,numeros);
            checkCpf(numeros);

            if(checkCpf(numeros)==false){
                Console.WriteLine("cpf falso");
                
            }
           


      
           

        }


    }
}
