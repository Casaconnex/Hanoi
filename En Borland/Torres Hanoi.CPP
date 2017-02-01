#include"stdlib.h"
#include"conio.h"
#include"string.h"
#include"math.h"
#include"stdio.h"
#define  lo 10


//int lo;
int a[lo],b[lo],c[lo];
int pos1, pos2, pos3;
int n;
int ant,sig;

void imprimir();
void inicializar();
void juego();
void mov1();
void mov2();
void mov3();

void main()
{
clrscr();
printf("digite el tamaño del vector: ");
scanf("%d",&n);
inicializar();
juego();
}

void imprimir(){
int i;
clrscr();
for(i=1;i<=n;i++)
{gotoxy(10,6+i);
 printf("%d",a[i]);
 gotoxy(14,6+i);
 printf("%d",b[i]);
 gotoxy(18,6+i);
 printf("%d",c[i]);
 }
 gotoxy(20,10);
 printf("pos1= %d",pos1);
 gotoxy(30,10);
 printf("pos2= %d",pos2);
 gotoxy(40,10);
 printf("pos3= %d",pos3);
 getch();

}

void inicializar()
{
  int i;

  for(i=1;i<=n; i++)
  {
   	a[i]=i;
      b[i]=0;
      c[i]=0;
   }
   pos1=1;
   pos2=n;
   pos3=n;
   ant=0;
   sig=0;

}

void juego()
{
   if (ant==0 && sig ==0)
   	mov1();
   else{
   	if (ant==1 && sig==2)
      	mov3();
      else if(ant==1 && sig==3)
      	mov2();
      else if(ant==2 && sig==3)
      	mov1();
      else if(ant==2 && sig==1)
      	mov3();
      else if(ant==3 && sig==1)
      	mov2();
      else if(ant==3 && sig==2)
      	mov1();

   }
   if(pos1==n && pos2==n && pos3==1)
   	{
      	gotoxy(10,22);
         printf("TERMINO EL JUEGO");
         getch();
         exit(0);
      }
   juego();
}
void mov1()
{

	int parn;
	int impar1;
   imprimir();
   if (pos1>=n)
   	pos1=n;

   if(pos1<=n)
   {
      parn=n%2;
      impar1=pos1%2;

      if(parn==impar1)
      {
       if(c[pos3]==0)
       {
         if(a[pos1]!=0){
       	c[pos3]=a[pos1];
         a[pos1]=0;
         pos1++;
         ant=1;
         sig=3;
         mov1();
         }
       }
       else
       {
       	if(c[pos3]>a[pos1])
         {
         if(a[pos1]!=0)
  			{
         pos3--;
       	c[pos3]=a[pos1];
         a[pos1]=0;
         pos1++;
         ant=1;
         sig=3;
         mov1();
         }}
       }
      }
     else
     {
       if(b[pos2]==0)
       {
         if(a[pos1]!=0)
  			{
       	b[pos2]=a[pos1];
         a[pos1]=0;
         pos1++;
         ant=1;
         sig=2;
         mov1();
         }
       }
       else
       {
       	if(b[pos2]>a[pos1])
         {
         if(a[pos1]!=0)
  			{
         pos2--;
       	b[pos2]=a[pos1];
         a[pos1]=0;
         pos1++;
         ant=1;
         sig=2;
         mov1();
         }
         }
       }
     }
   }

}

void mov2()
{

	int parn;
	int impar1;
   imprimir();
   if(pos2>=n)
   	pos2=n;
	if(pos2<=n)
   {

   	parn=n%2;
      impar1=(pos2)%2;

      if(parn==impar1)
      {
         if(c[pos3]==0)
         {
            if(b[pos2]!=0)
  			{
         	c[pos3]=b[pos2];
            b[pos2]=0;
            pos2++;
            ant=2;
            sig=3;
            mov2();
            }
         }
         else
         {
         	if(c[pos3]>b[pos2])
            {
              if(b[pos2]!=0)
  			{
              pos3--;
              c[pos3]=b[pos2];
              b[pos2]=0;
              pos2++;
              ant=2;
              sig=3;
              mov2();
            }
         }
         }}

         else
         {
            if(a[pos1]==0)
         {
            if(b[pos2]!=0)
  			{

         	a[pos1]=b[pos2];
            b[pos2]=0;
            pos2++;
            ant=2;
            sig=1;
            mov2();
         }}
         else
         {
         	if(a[pos1]>b[pos2])
            {
              if(b[pos2]!=0)
  			{
              pos1--;
              a[pos1]=b[pos2];
              b[pos2]=0;
              pos2++;
              ant=2;
              sig=1;
              mov2();
            }}
         }}}


}

void mov3()
{

	int parn;
	int impar1;
   imprimir();
   if(pos3>=n)
   	pos3=n;
	if(pos3<=n)
   {

   	parn=n%2;
      impar1=(pos3)%2;

      if(parn==impar1)
      {
    		if(b[pos2]==0)
				{
               if(c[pos3]!=0)
  			{
            	b[pos2]=c[pos3];
               c[pos3]=0;
            	pos3++;
               ant=3;
               sig=2;
               mov3();
               }
            }
         else
         {
         	if(b[pos2]>c[pos3])
            {
              if(c[pos3]!=0)
  			{
              pos2--;
              b[pos2]=c[pos3];
               c[pos3]=0;
            	pos3++;
               ant=3;
               sig=2;
               mov3();
          }  }
         }
      }
      else
      {
          if(a[pos1]==0)
				{
					if(c[pos3]!=0)
  			{
            	a[pos1]=c[pos3];
               c[pos3]=0;
            	pos3++;
               ant=3;
               sig=1;
               mov3();
            }}
         else
         {
         	if(a[pos1]>c[pos3])
            {
              if(c[pos3]!=0)
  			{
              pos1--;
              a[pos1]=c[pos3];
               c[pos3]=0;
            	pos3++;
               ant=3;
               sig=1;
               mov3();
            }}
         }
      }
   }
}
