# **V-Belt Drive Calculator and 3D model Generator - Visual Basic 6**

![animation](/img/main-animation.gif)

</br> Building this Software was a part of my Bachelor's thesis at the University of Nis, Faculty of Mechanical Engineering. Software was developed in **Visual Basic 6**.
</br>Thesis was: "Automation of the Calculation and Design of the V-Belt Drives". Besides calculations, I made an API connection with SolidWorks for generating a 3D model calculated pulley.
</br>Mentor for my Bachelor thesis was professor Dragan Milčić, PhD in Mechanical Engineering.

</br>Software features:
- V-Belt Pulley Calculation, including
    - Normal profile according to ISO 4184: 1992 / DIN 2215: 1998
    - Narrow profile according to DIN 7753
- Possibility to use stronger profiles after calculation 
- Generating calculation report in Word (.doc)
- Generating calculated 3D model of drive pulley and/or driven pulley in SolidWorks
- Generating 3D model of drive pulley and/or driven pulley independently of the calculation

---

## **Table of Contents**
- [Main algorithm](#main-algorithm)
- [Tables and Diagrams](#tables-and-diagrams)
- [Use of Wolfram Mathematica](#use-of-wolfram-mathematica)
    - [Finding interpolation polynomial of the Wrapping angle factor](#finding-interpolation-polynomial-of-the-wrapping-angle-factor)
    - [Determination of the nominal specific Pn power of one narrow belt](#determination-of-the-nominal-specific-pn-power-of-one-narrow-belt)
- [The Example of Final Calculation Report](#the-example-of-final-calculation-report)
- [The Example of generated 3D model in SolidWorks](#the-example-of-generated-3d-model-in-solidworks)
- [Pulley Drawings - Examples](#pulley-drawings---examples)

---

## **Main algorithm**
Main algorithm is shown on the picture below, it's written on Serbian language:
- Entry data:
    - Input power P [kW]
    - Input RPM n1 [min^-1]
    - Ratio i
    - Working conditions factor Ka

![algorithm](/img/algorithm.jpg)

---

## **Tables and Diagrams**
Calculation of the V-Belt drives requires using data from tables and diagrams. Good thing is that all of these data are implemented into the software so user don't have to worry about that.
For achieving this, I used:
- GetData Graph Digitizer and
- Wolfram Mathematica.

GetData Graph Digitizer is a program for digitizing graphs and plots. It is often necessary to obtain original (x,y) data from graphs, e.g. from scanned scientific plots, when data values are not available.
</br>Wolfram Mathematica is used for interpolation and fitting curves from data obtained from GetData Graph Digitizer. Examples are below.

## **Use of Wolfram Mathematica**
In this section I will represent only a few examples of 'transforming' data from tables, diagrams and plots into mathematical equations.

### **Finding interpolation polynomial of the Wrapping angle factor**
Here is an example how the values for the Wrapping angle factor are 'transformed' (fitted) into a single equation.
</br>Table 1. Wrapping angle factor

![cbeta](/img/cbeta.png)
By using a simple command in Mathematica:
```
Fitting = Fit[Lista, {1, x, x^2, x^3}, x];
```
We will get an interpolation polynomial:
![fitting](/img/fitting.png)
By having interpolation polynomial, we can calculate Wrapping angle factor for any angle.
In folder 'wolfram_mathematica_notebooks', whole notebook is available and also PDF.

### **Determination of the nominal specific Pn power of one narrow belt**
Here is a diagram for determining the nominal specific Pn power of one SPA (DIN 7753) narrow belt:

![spa](/img/spa.jpg)

</br> For every curve, we have to find an interpolation polynomial.
By using GetData Graph Digitizer we are getting a values Pn [kW] and RPM [min^-1]

```
Lista250 = {{100, 1.22}, {200, 2.3}, {300, 3.25}, {500, 5.16}, {700, 
    7}, {950, 9}, {1450, 12.58}, {2000, 15.8}, {2850, 19.21}};
Lista200 = {{111.8, 1}, {200, 1.68}, {300, 2.4}, {500, 3.84}, {700, 
    5.25}, {950, 6.83}, {1450, 9.8}, {2000, 12.52}, {2850, 15.8}};
Lista160 = {{156, 1}, {200, 1.24}, {300, 1.76}, {500, 2.78}, {700, 
    3.8}, {950, 5}, {1450, 7.2}, {2000, 9.3}, {2850, 11.81}};
Lista125 = {{240, 1}, {300, 1.24}, {500, 1.96}, {700, 2.65}, {950, 
    3.43}, {1450, 4.86}, {2000, 6.21}, {2850, 7.94}};
Lista112 = {{300, 1}, {500, 1.58}, {700, 2.15}, {950, 2.8}, {1450, 
    4}, {2000, 5.08}, {2850, 6.56}};
Lista90 = {{530, 1}, {700, 1.31}, {950, 1.7}, {1450, 2.4}, {2000, 
    3.03}, {2850, 3.8}};

```
Fitting:
```
d250[x_] = Fit[Lista250, {1, x, x^2}, x];
Print["d250=", d250[x]]
d200[x_] = Fit[Lista200, {1, x, x^2}, x];
Print["d200=", d200[x]]
d160[x_] = Fit[Lista160, {1, x, x^2}, x];
Print["d160=", d160[x]]
d125[x_] = Fit[Lista125, {1, x, x^2}, x];
Print["d125=", d125[x]]
d112[x_] = Fit[Lista112, {1, x, x^2}, x];
Print["d112=", d112[x]]
d90[x_] = Fit[Lista90, {1, x, x^2}, x];
Print["d90=", d90[x]]
```
Here is plotted interpolation polynomial:
![spa-fitting](/img/spa-fitting.png)

As the previous one, also this Mathematica notebook and PDF is available in folder 'wolfram_mathematica_notebooks'.

---

## **The Example of Final Calculation Report**
At the end of the Calculation, we have the possibility to generate One-Page Report in Microsoft Word (.doc).
</br>It is on Serbian language. Here is how it looks:

```
Mašinski fakultet Univerziteta u Nišu
Katedra za Mašinske konstrukcije, razvoj i inženjering
Obradio: Nikola Lazić
Datum: 18.1.2018
Proračun remena normalne širine
Ulazni podaci:
Nominalna snaga na ulazu	P1=18.5 kW
Broj obrtaja na ulazu	n1=1450 min-1
Prenosni odnos	i=2
Faktor radnih uslova	KA=1.3

Proračun:
Prečnik pogonske remenice	d1=160 mm
Prečnik gonjene remenice	d2=315 mm
Računska vrednost osnog rastojanja	a=700 mm
Ugao nagiba ogranka	α=6.356O
Obvojni ugao	β1=167.287O
Računska vrednost dužine kaiša	Ldr=2114.717 mm
Korekcija dužine kaiša	ΔL=40 mm
Stvarno osno rastojanje	as=642.26 mm
Nominalna specifična snaga	PN=6.309 kW
Dodatna snaga po remenu	ΔP=0.658 kW
Faktor dužine remena	CL=0.979
Faktor obvojnog ugla	Cβ=0.967
Računska vrednost broja žlebova	zr=3.645
Obimna brzina	ν1=12.147 m/s
Maksimalna obimna brzina	νmax=30 m/s
Učestanost savijanja	fs=11.909 s-1
Dozvoljena učestanost savijanja	fsdoz=80 s-1

Rezultati:
Profil kaiša	Normalni
Tip kaiša	B/17 - ISO 4184: 1992 / DIN 2215: 1998
Standardna vrednost dužine kaiša	Ld=2040 mm
Broj žlebova	z=4
```
---

## **The Example of generated 3D model in SolidWorks**
One of the feature is generating a 3D model of calculated drive/driven pulley or generating 3D model of drive/driven pulley independently of the calculation. 
![cad-generating](/img/cad-animation.gif)

</br>Generated 3D model in SolidWorks:
![3d model in SolidWorks](/img/3d-sw.png)

After 3D model generation, user have to design the body of the pulley.
</br>Here is some examples how Pulley should look at the end:
![pulley-1](/img/pulley-1.png)
![pulley-2](/img/pulley-2.png)

---

## **Pulley Drawings - Examples**
![drawing-1](/img/drg-1.png)
![drawing-2](/img/drg-2.png)
