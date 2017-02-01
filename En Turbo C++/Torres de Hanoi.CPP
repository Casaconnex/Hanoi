#include <stdio.h>
#include <conio.h>
#include <stdlib.h>
#include <dos.h>
#include <graphics.h>
#include <math.h>
#include <string.h>



int torre1[30], torre2[30], torre3[30];
int x1, x2, x3;
int movoriginal, movfinal, mov;

void inicio_grafico();
void mostrar(int cant);
void llenar(int cant);
int pregunta(char * msg);

void hanoi(int cant);
void hanoi_T1(int cant);
void hanoi_T2(int cant);
void hanoi_T3(int cant);



void main(){
	int canti;
   clrscr();
   inicio_grafico();
   setcolor(RED);
   canti=pregunta("INGRESE LA CANTIDAD DE ELEMENTOS DE LA TORRE 1");
   llenar(canti);
   hanoi(canti);
   getch();
   sound(0);
}


void hanoi(int cant){

	setfillstyle(1,11);
	bar(0,0,680,800);
   if(movoriginal==1)
	   {
      	if(movfinal==1)
	         hanoi_T1(cant);
         else if(movfinal==2)
         	hanoi_T3(cant);
         else if(movfinal==3)
         	hanoi_T2(cant);
   	}
   else if(movoriginal==2)
	{
			if(movfinal==1)
		hanoi_T3(cant);
	 else if(movfinal==3)
		hanoi_T1(cant);
      }
   else if(movoriginal==3)
	{
			if(movfinal==1)
		hanoi_T2(cant);
	 else if(movfinal==2)
		hanoi_T1(cant);
      }
      if(x1==cant && x2==cant && x3==1)
	{

	 char *msg;
	 int y;
	 msg="JUEGO TERMINADO";
	 settextstyle(7,0,5);
	 y=1;
	 do{
		 randomize();
		 setcolor(rand()%y);
		 outtextxy((680-textwidth(msg))/2,400,msg);
		 delay(50);
		 y++;
	 }while(y<=100);
	 getch();
	 exit(0);
      }

      hanoi(cant);
}

void hanoi_T1(int cant){
   int par, par_torre1;
   setfillstyle(1,128);
   bar(5,105,cant*20+5,cant*10+115);
   setfillstyle(1,14);
   bar(0,100,cant*20,cant*10+110);
   mostrar(cant);
   par=cant%2;
   par_torre1=x1%2;

   if(torre1[cant]!=0)
	{
       	if(par==par_torre1)
		{
		if(torre3[cant]==0 && cant==x3)
		{
		     torre3[x3]=torre1[x1];
		     torre1[x1]=0;
		     x1++;
		     if(x1>cant)
			x1=cant;
		     movoriginal=1;
		     movfinal=3;
		     hanoi_T1(cant);
		  }
	       else
		{
			if(torre3[x3]>torre1[x1])
			{
			   x3--;
			   torre3[x3]=torre1[x1];
				torre1[x1]=0;
		                     x1++;
      		               if(x1>cant)
				x1=cant;
			      movoriginal=1;
		                     movfinal=3;
			       hanoi_T1(cant);
                        }
                  }
            }
	 else
         	{
		if(torre2[cant]==0 && cant==x2)
               	{
							torre2[x2]=torre1[x1];
		     torre1[x1]=0;
                     x1++;
		     if(x1>cant)
                     	x1=cant;
                     movoriginal=1;
		     movfinal=2;
                     hanoi_T1(cant);
                  }
	       else
		{
                  	if(torre2[x2]>torre1[x1])
			{
                           x2--;
									torre2[x2]=torre1[x1];
                     		torre1[x1]=0;
				     x1++;
      		               if(x1>cant)
				x1=cant;
               		      movoriginal=1;
				     movfinal=2;
			       hanoi_T1(cant);
                        }
		  }
            }
      }
}


void hanoi_T2(int cant){
   int par, par_torre2;
   setfillstyle(1,128);
   bar(225,105,225+cant*20,cant*10+115);
   setfillstyle(1,14);
   bar(220,100,220+cant*20,cant*10+110);

   mostrar(cant);
   par=cant%2;
   par_torre2=x2%2;

   if(torre2[cant]!=0)
	{
       	if(par==par_torre2)
		{
		if(torre3[cant]==0 && cant==x3)
		{
							torre3[x3]=torre2[x2];
		     torre2[x2]=0;
		     x2++;
		     if(x2>cant)
			x2=cant;
		     movoriginal=2;
		     movfinal=3;
                     hanoi_T2(cant);
		  }
	       else
		{
                  	if(torre3[x3]>torre2[x2])
			{
                           x3--;
									torre3[x3]=torre2[x2];
				torre2[x2]=0;
		                     x2++;
			       if(x2>cant)
				x2=cant;
			      movoriginal=2;
				     movfinal=3;
      		               hanoi_T2(cant);
                        }
                  }
            }
	 else
		{
            	if(torre1[cant]==0 && cant==x1)
		{
							torre1[x1]=torre2[x2];
		     torre2[x2]=0;
                     x2++;
		     if(x2>cant)
                     	x2=cant;
		     movoriginal=2;
		     movfinal=1;
                     hanoi_T2(cant);
		  }
	       else
		{
			if(torre1[x1]>torre2[x2])
                     	{
                           x1--;
									torre1[x1]=torre2[x2];
                     		torre2[x2]=0;
				     x2++;
			       if(x2>cant)
            	      	   	x2=cant;
			      movoriginal=2;
				     movfinal=1;
			       hanoi_T2(cant);
                        }
		  }
            }
      }
}

void hanoi_T3(int cant){
   int par, par_torre3;
   setfillstyle(1,128);
   bar(425,105,cant*20+425,cant*10+115);
   setfillstyle(1,14);
   bar(420,100,420+cant*20,cant*10+110);
   mostrar(cant);
   par=cant%2;
   par_torre3=x3%2;

   if(torre3[cant]!=0)
	{
	if(par==par_torre3)
		{
            	if(torre2[cant]==0 && cant==x2)
		{
							torre2[x2]=torre3[x3];
		     torre3[x3]=0;
		     x3++;
                     if(x3>cant)
			x3=cant;
                     movoriginal=3;
                     movfinal=2;
		     hanoi_T3(cant);
                  }
               else
		{
                  	if(torre2[x2]>torre3[x3])
                     	{
			   x2--;
									torre2[x2]=torre3[x3];
                     		torre3[x3]=0;
				     x3++;
      		               if(x3>cant)
            	      	   	x3=cant;
			      movoriginal=3;
				     movfinal=2;
			       hanoi_T3(cant);
			}
                  }
	    }
         else
         	{
		if(torre1[cant]==0 && cant==x1)
               	{
							torre1[x1]=torre3[x3];
		     torre3[x3]=0;
                     x3++;
                     if(x3>cant)
			x3=cant;
                     movoriginal=3;
                     movfinal=1;
		     hanoi_T3(cant);
                  }
               else
		{
			if(torre1[x1]>torre3[x3])
			{
			   x1--;
									torre1[x1]=torre3[x3];
				torre3[x3]=0;
		                     x3++;
      		               if(x3>cant)
				x3=cant;
               		      movoriginal=3;
		                     movfinal=1;
			       hanoi_T3(cant);
                        }
                  }
	    }
      }
}


void mostrar(int cant){


   float j,a,b,c,d;

   setfillstyle(1,BLACK);
   a=(cant*20-10)/2;
   b=105;
   c=a+10;
   d= b+cant*10+10;
   bar(a,b,c,d);
   setcolor(BLACK);
   line(0,d,cant*20,d);

   a= a+220;
   c=a+10;
   bar(a,b,c,d);
   setcolor(BLACK);
   line(220,d,220+cant*20,d);

   a=a+200;
   c=a+10;
   bar(a,b,c,d);
   setcolor(BLACK);
   line(420,d,420+cant*20,d);

   for(int i=1;i<=cant;i++)
	{
	if(torre1[i]!=0)
		{
	       setfillstyle(1,torre1[i]);
	       a=(((cant-torre1[i]))*10);
	       b=10*i+100;
	       c=cant*20-a;
	       d=b+10;
	       bar(a,b,c,d);
	    }
	 if(torre2[i]!=0)
		{
		setfillstyle(1,torre2[i]);
		a=220+(((cant-torre2[i]))*10);
		b=10*i+100;
		c=a+torre2[i]*20;
		d=b+10;
		bar(a,b,c,d);
	    }
	 if(torre3[i]!=0)
		{
		setfillstyle(1,torre3[i]);
		a=420+(((cant-torre3[i]))*10);
		b=10*i+100;
		c=a+torre3[i]*20;
		d=b+10;
		bar(a,b,c,d);
	    }
      }
   delay(500);
   sound(0);
}


void llenar(int cant){
	for(int i=1;i<=cant;i++)
	{
	torre1[i]=i;
	 torre2[i]=0;
	 torre3[i]=0;
      }
   x1=1;
   x2=cant;
   x3=cant;
   movoriginal=1;
   movfinal=1;

}

void inicio_grafico()
{
   int controlador, modo;
   controlador=DETECT;
   modo=2;
   initgraph(&controlador,&modo, "c:\\tc\\bgi");
   setfillstyle(1,BLUE);
   bar(0,0,640,480);
}



int pregunta(char *msg)
{
	int t,a,b,c,d;
	char *fr;
	a=(640-textwidth(msg))/2 - 10;
	b= (480-textheight(msg))/2-10;
	c= a+ textwidth(msg)+10;
	d=b+textheight(msg)+10;
	setfillstyle(1,BLACK);
	bar(a+5,b+5,c+5,d+5);
	setfillstyle(1,LIGHTGRAY);
	bar(a,b,c,d);
	setcolor(GREEN);
	outtextxy(a+5,b+5,msg);
	setfillstyle(1,LIGHTGREEN);
	bar(300,300,340,310);
	int pos=0;
	do{
		t=getch();
		if(t!=13){
		if(t<48 || t>57){
			t=0;
			}
		else{
		fr[pos]=t;
		fr[pos+1]='\0';
		setcolor(RED);
		if(pos>1)
			break;
		outtextxy(300,300,fr);
		pos++;}}
	}while(t!=13);
	int i=pos-1;
	t=0;

	while(i>=0)
		{
		t+=pow(10,strlen(fr)-i-1)*(fr[i]-48);
		i--;
		}
	if(t<=15 && t>0)
		return t;
	else
		return pregunta(msg);
}