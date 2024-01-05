#include "MiniExcel.h"


int main()
{
	Mini_Excel x;
	x.use_Excel();

}
//int main()
//{
//	int _r=100, _c=115;
//	while (true)
//	{
//		gotoRowCol(_r, _c);
//		cout << "                                        ";
//		gotoRowCol(0, 0);
//		cout << "                                               ";
//		gotoRowCol(0, 0);
//		cout<< "1st enter row then col where you want to go: ";
//		cin >> _r >> _c;
//		gotoRowCol(_r, _c);
//		cout << "This is the row and col you entered";
//		_getch();
//	}
//}



//int main()
//{
//    while (true)
//    {
//        int x = 0;
//        char y;
//        y = _getch();
//        x = y;
//        cout << "Ascii value: " << x << " and its char: " << y << endl;
//    }
//    return 0;
//}


/*
++++++++++++++     Keys and their relative commands for using My Mini_Excel   ++++++++++++++   
Note: The character commands support caps lock of and on (both)  except for commands with signs like the front slash
But it is recommended that you keep the 'caps look of'.
						
  Command :/ -> (front slash) :   for selecting and deselecting a cell .  
  Command: D for deleting -> (C for column/ R for row)  -> [(R for right/ L for left) / (A for above / B for below)]
  Command: I for insert....   (next commands are same as above)
  Command: C for cell  ->  S for alignmnet style / k for colour of cell
  
  Command: M for going i




*/
