#pragma once
#include <iostream>
#include <string.h>
#include "goto-row-column and utility functions.h"
using namespace std;

static int cell_row_dim = 5;
static int cell_col_dim = 13;
static char symbol = '*';

class Mini_Excel
{
	//+++++++++++     Private  friend classes section
	
	//---- cell class
	class cell
	{
	
		string data; // data stored 
		cell* up;
		cell* down;
		cell* left;
		cell* right;
		int text_clr;
		char cell_alignment; // c for centre, r for right, l for left
		//int _r, _c; // cell knows in which column and row he is present in 
		friend Mini_Excel; // Mini_Excel is a friend class and it can use cell's private attributes
		friend class XL_Iterator;
	//public:
		cell(string s = "", cell* u = nullptr, cell* d = nullptr, cell* l = nullptr, cell* r = nullptr, char cell_alignment = 'c') //,int row=0,int col=0
		{
			data = s;
			up = u;
			down = d;
			left = l;
			right = r;
			text_clr = WHITE;
			cell_alignment = cell_alignment;  
			//_r = row; _c = col;
		}
		cell(const cell& x)
		{
			data = x.data;
			up = x.up;
			down = x.down;
			left = x.left;
			right = x.right;
			text_clr = x.text_clr;
			cell_alignment = x.cell_alignment;
		}
		bool operator ==(const cell& x)
		{
			return &(*this) == &x;
		}
		void swap_cells(cell& cl)
		{
			cell decoy(cl);
			cl.data = this->data;
			cl.left = this->left;
			cl.right = this->right;
			cl.up = this->up;
			cl.down = this->down;
			cl.text_clr = this->text_clr;
			cl.cell_alignment = this->cell_alignment;

			this->data = decoy.data;
			this->up = decoy.up;
			this->down = decoy.down;
			this->left = decoy.left;
			this->right = decoy.right;
			this->text_clr = decoy.text_clr;
		}

		/*void tell_row_col(int row,int col)
		{
			this->_r = row;
			this->_c = col;
		}*/
		//~cell() // cell destructor
		//{
		//	up = nullptr;
		//	down = nullptr;
		//	left = nullptr;
		//	right = nullptr;

		//}

	};

	//--- Iterator class of MiniExcel
	class XL_Iterator
	{
		cell* current;
		friend Mini_Excel; // excel can use private attributs of XL_Iterator
		friend cell;
		XL_Iterator()
		{
			current = nullptr;
		}

		XL_Iterator(cell* c)
		{
			current = c;
		}

		XL_Iterator(const XL_Iterator& x)
		{
			current = x.current;
		}
	public:
		void set_pointer(cell* c)
		{
			current = c;
		}
		bool operator ==(const XL_Iterator& x)
		{
			current = x.current; // just checking the addresses of the cells these are pointing at
		}
		void operator= (const XL_Iterator& x)
		{
			current = x.current;
		}

		//_____ pre increament and decrement
		// ++++++++++++++++ using for moving dowm and above in rows (horizontally)
		void operator++ () 
		{
			current = current->down; // moving down the row
		}
		void operator-- ()
		{
			current = current->up; // moving up
		}

		//____ Post increment and decrement
		//+++++++++++  usign for moving in verticle dimension (from left to right and vice versa)

		void operator ++(int)
		{
			current = current->right; // moving right
		}

		void operator --  (int)
		{
			current = current->left; // moving left
		}

	};
	
	//++++++++++++      Private variable section

	int current_row, current_col; // row and column of the current cell
	cell head; // this is the 1st cell of the whole 2d matrix
	cell* head_pointer;
	XL_Iterator current_cell;
	XL_Iterator range_start;
	XL_Iterator range_end;
	/*cell* current_cell;
	cell* range_start;
	cell* range_end;*/
	
	// dimension of total cell grid on the screen (initially 5x5)
	int grid_row_dim;
	int grid_col_dim;
	bool is_a_cell_selected;

	//+++++++++++     private functions section
	void InsertCellToRight(cell*& cur) // passed the current cell
	{						// cur = current_cell
		cell* decoy = new cell;
		//cell* cur_temp = cur;

		// current cell is in the last column
		if (cur->right == nullptr)
		{		// a new cell with no up and down only left pointed to current cell
			cur->right = decoy;
			decoy->left = cur;
		}// current cell already has a cell on its right
		else if (cur->right != nullptr)
		{
			decoy->right = cur->right;
			cur->right->left = decoy;
			decoy->left = cur;
			cur->right = decoy;

		}
		// for now we are settign up and down of new cell to null 
		// and we will change it once we have constructed the whole column
		decoy->up = nullptr;
		decoy->down = nullptr;
	}

	void InsertCellToLeft(cell* cur)
	{
		cell* decoy = new cell;
		//cell* cur_temp = cur;

		// current cell is in the 1st column
		if (cur->left == nullptr)
		{		// a new cell with no up and down only right pointed to current cell
			cur->left = decoy;
			decoy->right = cur;
		}// current cell already has a cell on its left
		else if (cur->right != nullptr)
		{
			decoy->right = cur;
			cur->left->right = decoy;
			decoy->left = cur->left;
			cur->left = decoy;

		}
		// for now we are settign up and down of new cell to null 
		// and we will change it once we have constructed the whole column	
		decoy->up = nullptr;
		decoy->down = nullptr;
	}

	void InsertCellUp(cell* cur)  // cur = current cell
	{
		cell* decoy = new cell; // a new cell memory
		// if the current cell is on top 
		if (cur->up == nullptr)
		{
			cur->up = decoy;
			decoy->down = cur;
		} // agr centre ma kahin ha
		else
		{
			cur->up->down = decoy;
			decoy->up = cur->up;
			decoy->down = cur;
			cur->up = decoy;
		}
		// initially setting left and right if the new cell equal to null
		decoy->left = nullptr;
		decoy->right = nullptr;

	}

	void InsertCellDown(cell* cur)  // cur = current cell
	{
		cell* decoy = new cell; // a new cell memory
		// if the current cell is on bottom 
		if (cur->down == nullptr)
		{
			cur->down = decoy;
			decoy->up = cur;
		} // agr centre ma kahin ha
		else
		{
			cur->down->up = decoy;
			decoy->up = cur;
			decoy->down = cur->down;
			cur->down = decoy;
		}
		// initially setting left and right if the new cell equal to null
		decoy->left = nullptr;
		decoy->right = nullptr;

	}



	
			//++++++   Insertion at right  ++++++++
	void InsertColumnAtRight()
	{
		XL_Iterator decoy(this->current_cell); // saving the cell at which we are right now
		// going at the start of the column in which our current cell is present
		while (decoy.current->up != nullptr)
		{
			--decoy; // going up
		}
		XL_Iterator top(decoy);// top cell of the column ko save kar rha hon

 		while (decoy.current != nullptr) // hum na last cell ka aga bhi insert karna ha or us sa aga ya null ho jai ga
		{
			InsertCellToRight(decoy.current);
			++decoy; // moving dowm in column
		}

		// now changing the up and down pointers of the new column we just inserted
		while (top.current != nullptr)
		{
			
			if (top.current->up == nullptr && top.current->down == nullptr)
			{
				//continue;
				//do nothing
			} 
			// if we are on the top of column
			else if (top.current->up == nullptr)
			{
				top.current->right->down = top.current->down->right;
			}  // if we are at the end of column
			else if (top.current->down == nullptr)
			{
				top.current->right->up = top.current->up->right;
			} // if we are in between
			else
			{
				top.current->right->up = top.current->up->right;
				top.current->right->down = top.current->down->right;
			}
			top.current = top.current->down;
		}
		this->grid_col_dim++;
	}

			//++++++++++   Insertion at left  +++++++++++++
	
	void InsertColumnAtLeft()
	{
		bool fl = false;
		XL_Iterator decoy(this->current_cell); // saving the cell at which we are right now
		// going at the start of the column in which our current cell is present
		while (decoy.current->up != nullptr)
		{
			--decoy; // going up
		}
		XL_Iterator top(decoy);// top cell of the column ko save kar rha hon
		if (*(top.current) == *this->head_pointer)  // agr us column ka left pa ak new column insert karna ha jis ma head majood ha too head ko bhi dobara assign karna para ga 
		{
			fl = true;
		}
		while (decoy.current != nullptr) // hum na last cell ka aga bhi insert karna ha or us sa aga ya null ho jai ga
		{
			InsertCellToLeft(decoy.current);
			++decoy; // moving dowm in column
		}
		if (fl)
		{    // reassigning of head if it is pushed to right
			this->head_pointer = top.current->left;
		}
		// now changing the up and down pointers of the new column we just inserted
		while (top.current != nullptr)
		{
			if (top.current->up == nullptr && top.current->down == nullptr)
			{
				//continue;
				// do nothing
			}
			// if we are on the top of column
			else if (top.current->up == nullptr)
			{
				top.current->left->down = top.current->down->left;
				
			}  // if we are at the end of column
			else if (top.current->down == nullptr)
			{
				top.current->left->up = top.current->up->left;

			} // if we are in between
			else
			{
				top.current->left->up = top.current->up->left;
				top.current->left->down = top.current->down->left;
			}
			top.current = top.current->down;
		}
		this->current_col++; // since current cell ka picha ak column add hoa ha is lia us ka column number ak aga chala gia ha
		this->grid_col_dim++;
	}

	//++++++++   Insertion on above
	
	void InsertRowAbove()
	{
		bool fl=false;
		XL_Iterator decoy(this->current_cell); // saving the cell at which we are right now
		// going at the start of the row in which our current cell is present
		while (decoy.current->left != nullptr)
		{
			decoy--; // going left
		}
		XL_Iterator left_most(decoy);// left cell of the row ko save kar rha hon
		if (*(left_most.current) == *this->head_pointer)
		{
			fl = true;
		}
		   ///++++++    inserting the cells
		while (decoy.current != nullptr) // hum na last cell ka uper bhi insert karna ha or us sa aga ya null ho jai ga
		{
			InsertCellUp(decoy.current);
			decoy++; // moving right in row
		}
		if (fl)  // if cell is placed above the head pointr then reassign the head pointer to it
		{
			head_pointer = left_most.current->up;
		}

		// now changing the configuration of the cells we just inserted
		while (left_most.current != nullptr)
		{
			if (left_most.current->left == nullptr && left_most.current->right == nullptr)
			{
				// do nothing
			}
			// we are at left most
			else if (left_most.current->left == nullptr /*&& left_most.current->right != nullptr*/)
			{
				left_most.current->up->right = left_most.current->right->up;
			}
			// we are at right most
			else if (/*left_most.current->left != nullptr &&*/ left_most.current->right == nullptr)
			{
				left_most.current->up->left = left_most.current->left->up;
			} // we are in centre
			else
			{
				left_most.current->up->left = left_most.current->left->up;
				left_most.current->up->right = left_most.current->right->up;
			}
			left_most.current = left_most.current->right; // going from left to right
		}
		this->current_row++; // koinka current ka uper ak row insert hoi ha too wo ak row nicha chala gia ha
		this->grid_row_dim++;
	}

	//++++++++   Insertion at below
	

	void InsertRowBelow()
	{
		XL_Iterator decoy(this->current_cell); // saving the cell at which we are right now
		// going at the start of the row in which our current cell is present
		while (decoy.current->left != nullptr)
		{
			decoy--; // going left
		}
		XL_Iterator left_most(decoy);// left cell of the row ko save kar rha hon

		while (decoy.current != nullptr) // hum na last cell ka uper bhi insert karna ha or us sa aga ya null ho jai ga
		{
			InsertCellDown(decoy.current);
			decoy++; // moving right in row
		}

		// now changing the configuration of the cells we just inserted
		while (left_most.current != nullptr)
		{
			if (left_most.current->left == nullptr && left_most.current->right == nullptr)
			{
				//continue;
				// do nothing
			}
			// we are at left most
			else if (left_most.current->left == nullptr /*&& left_most.current->right != nullptr*/)
			{
				left_most.current->down->right = left_most.current->right->down;
			}
			// we are at right most
			else if (/*left_most.current->left != nullptr &&*/ left_most.current->right == nullptr)
			{
				left_most.current->down->left = left_most.current->left->down;
			} // we are in centre
			else
			{
				left_most.current->down->left = left_most.current->left->down;
				left_most.current->down->right = left_most.current->right->down;
			}
			left_most.current = left_most.current->right;
		}
		this->grid_row_dim++; // increasing the dimesion record if the grid
	}
	

	//++++++++++  ( priting funtions ) ++++++++++++++
	

	// Thses 2 functions work together to print the data in cell
	
	void clear_data_in_cell(int which_row, int which_col)
	{
		gotoRowCol((which_row * cell_row_dim) + 2, (which_col * cell_col_dim) + 2);
		cout << "          ";
	}

	void print_data_in_cell(cell* c,int which_row, int which_col)
	{
		clear_data_in_cell(which_row,which_col);
		if (c->cell_alignment == 'l' || c->cell_alignment == 'L')  // left allignmnet
		{
			gotoRowCol((which_row * cell_row_dim) + 2, (which_col * cell_col_dim) + 2);
		}
		else if (c->cell_alignment == 'r' || c->cell_alignment == 'R') // right allignment
		{
			gotoRowCol((which_row * cell_row_dim) + 2, (which_col * cell_col_dim) + 7);
		}
		else // centre allignment
		{
			gotoRowCol((which_row * cell_row_dim) + 2, (which_col * cell_col_dim) + 5);
		}
		SetClr(c->text_clr);
		for (int i = 0; i < c->data.size(); i++)
		{
			cout << c->data[i]; // max it will contain 3 characters
		}
		SetClr(15); // back to white
	}
	

	// single cell printer
	void print_cell(cell* which_cell,int clr, int cur_row, int cur_col)
	{
		SetClr(clr);
		int strt_r = cur_row * cell_row_dim, strt_c = cur_col * cell_col_dim;
		gotoRowCol(cur_row * cell_row_dim, cur_col * cell_col_dim);
		for (int i = 0; i < cell_row_dim; i++)
		{
			for (int j = 0; j < cell_col_dim; j++)
			{
				gotoRowCol(strt_r + i, strt_c + j);
				if (i == 0 || j == 0 || i == cell_row_dim - 1 || j == cell_col_dim - 1)
				{
					cout << symbol;
				}
				
			}
		}
		SetClr(15); 
		print_data_in_cell(which_cell,cur_row,cur_col);
	}

	//+++++  this prints the whole excel sheet 
	void whole_sheet_printer(cell *head_poniter,int grid_total_row, int grid_total_col) // making no copy
	{
		
		//++++++++ 1st printing the function buttons 

		int buttons_start_col = 13;
		//gotoRowCol(5 * cell_row_dim, buttons_start_col);
		//____ printing the sum button
		cell sum_button("SUM");
		this->print_cell(&sum_button, WHITE, 2, buttons_start_col);
		//____  cout button
		cell count_button("COUNT");
		this->print_cell(&count_button, WHITE, 4, buttons_start_col);
		//____  Average button  
		cell AVG_button("AVG");
		this->print_cell(&AVG_button, WHITE, 6, buttons_start_col);
		//____  Min button  
		cell MIN_button("MIN");
		this->print_cell(&MIN_button, WHITE, 8, buttons_start_col);
		//____  Max button  
		cell MAX_button("MAX");
		this->print_cell(&MAX_button, WHITE, 10, buttons_start_col);

		//++++ Now Printing the grid
		//  This is going to be a zig zag printing 

		XL_Iterator printing_head(this->head_pointer); // iterator at the start of the grid
		for (int i = 0; i < grid_total_row; i++)
		{
			if (i % 2 == 0) // if i is even
			{
				for (int j = 0; j < grid_total_col; j++)
				{                  
					print_cell(printing_head.current,WHITE,i, j);
					if (j < grid_total_col - 1) // because if we have 2 cells then we will prin for 0 and 1 if we ++ after that the printing head will become null
					{
						printing_head++; // moving from most left to most right
					}
					
				}
				++printing_head; // moving one row down
			}
			else
			{
				for (int j = grid_total_col-1; j >= 0; j--)
				{
					print_cell(printing_head.current, WHITE, i, j);
					if (j > 0)
					{
						printing_head--; // moving from most right to most left
					}
				}
				++printing_head; // moving one row down
			}
			// Note: at the end the printing head iterator will become null
		}


	}

	void current_cell_printer(cell* curr_cell ,int cur_cell_row, int cur_cell_col)
	{
		this->print_cell(curr_cell, RED, cur_cell_row, cur_cell_col);
	}

	//++++  function for moving the current_cell in the direction  of the arrow key pressed
	void current_cell_mover(char arrow_key)
	{
		if (arrow_key == 'H') // up arrow key pressed
		{
			if (this->current_cell.current->up != nullptr)
			{
				print_cell(current_cell.current, WHITE, current_row, current_col);
				this->current_cell.current = this->current_cell.current->up;
				this->current_row--;
			}
		}
		else if (arrow_key == 'P') // down arrow key pressed
		{
			if (this->current_cell.current->down != nullptr)
			{
				print_cell(current_cell.current, WHITE, current_row, current_col);
				this->current_cell.current = this->current_cell.current->down;
				this->current_row++;
			}
		}
		else if (arrow_key == 'M') // right arrow key pressed
		{
			if (this->current_cell.current->right != nullptr)
			{
				print_cell(current_cell.current, WHITE, current_row, current_col);
				this->current_cell.current = this->current_cell.current->right;
				this->current_col++;
			}
		}
		else if (arrow_key == 'K') // left arrow key pressed
		{
			if (this->current_cell.current->left != nullptr)
			{
				print_cell(current_cell.current, WHITE, current_row, current_col);
				this->current_cell.current = this->current_cell.current->left;
				this->current_col--;
			}
		}
	}
	
	//       function for checking wheather the data about to be entered in the cell is valid 
	//		 (if there is already an int present in there then you cant put letters)
	//       and also is there space to enter it (max limit is 10)
	bool is_input_data_valid(cell* c,char x)
	{
		if (int(x) == -32)
		{
			x = _getch(); // if we have pressed arrow keys then we have to also waste its both characters
			return false;
		}

		if (int(x) == 8) // backspace char 
		{
			return true;
		}
		
		if (  !( (int(x) >= 48 && int(x) <= 57) || ((x >= 'A' || x >= 'a') && (x <= 'z' || x <= 'Z')) || x=='.')) // agr words, alpha bets or decimal ka alava ha too enter nhi karna
		{
			return false; // agr integers or alpha bets ka alava koi chz ha too wo wsa hi enter nhi ho gi
		}
		// agr abhi cell ha hi empty too direct put kar doo
		if (c->data.empty())
		{
			return true;
		}
		else if (c->data.size()==4) // agr wo limit pa ha too put nhi karna
		{
			return false;
		}
		else if (x == '.') // agr input decimal aya ha or pahla bhi ak decimal ha too usa new nahin dalna chia
		{
			for (int k{}; k < c->data.size(); k++)
			{
				if (c->data[k] == '.' || ((c->data[0] >= 'A' || c->data[0] >= 'a') && (c->data[0] <= 'z' || c->data[0] <= 'Z')))
				{
					return false;
				}
			}
		}
		else if ((int(c->data[0])>=48 && int(c->data[0]) <= 57) || c->data[0]=='.') // agr string ka pahla char number ha ya phir decimal ha
		{
			if ((x >= 'A' || x >= 'a') && (x <= 'z' || x <= 'Z')) // or jo char put karna lga han wo alphabet ha
			{
				// then it cannot be done
				return false;
			}
			
		}
		else if ((c->data[0] >= 'A' || c->data[0] >= 'a') && (c->data[0] <= 'z' || c->data[0] <= 'Z')) // agr string ka pahla char alphabet ha
		{
			if ((int(x) >= 48 && int(x) <= 57) || x=='.') // or tum number put karna lag ho ya decimal wo bhi allow nhi ha
			{
				return false;
			}
		}
		return true;
	}
	//____  function for entring data in a cell
	void enter_and_display_data_in_cell(cell* current_cell,int current_row,int current_col,char x)
	{
		if (int(x) != 8) // if x is not equal to backspace
		{
			current_cell->data.push_back(x);
		}
		else
		{
			if (!current_cell->data.empty()) // agr backspace press hoa ha or data ha cell ma too ak char delete kar do
			{
				current_cell->data.pop_back();
			}
		}
		print_data_in_cell(current_cell, current_row, current_col);
	}

//+++++++++++++       functions for deleting row and column
	
	//-- deleting columns
	// right
	void delete_right_cell(cell* dl)
	{
		// agr right pa koi cell hi nhi ha too kia delete karain
		if (dl->right == nullptr)
		{
			return; 
		} // agr right pa cell to ha mgr us ka right pa kuch nhi ha
		else if (dl->right->right == nullptr)
		{
			delete[]dl->right;
			dl->right = nullptr;
		} // last case ka jis cell ko delete karna wla han us ka right pa bhi cell ha
		else
		{
			cell* decoy = dl->right; // cell about to be deleted
			dl->right->right->left = dl;
			dl->right = dl->right->right;
			delete[]decoy;
		}

	}

	void delete_right_column()
	{
		XL_Iterator decoy(this->current_cell); // saving the cell at which we are right now
		if (this->current_cell.current->right!=nullptr) // because agr right pa koi cell nhi to koi column delete nhi ho ga or grid chota nhi ho ga
		{
			this->grid_col_dim--;
		}

		// going at the start of the column in which our current cell is present
		while (decoy.current->up != nullptr)
		{
			--decoy; // going up
		}
		XL_Iterator top(decoy);// top cell of the column ko save kar rha hon

		while (decoy.current != nullptr)
		{
			this->delete_right_cell(decoy.current);
			++decoy; // going down
		}
	}
	
	// left
	void deleting_left_cell(cell* dl)
	{
		if (dl->left == nullptr)
		{
			return;
		}
		else if (dl->left->left==nullptr)
		{
			delete[]dl->left;
			dl->left == nullptr;
		}
		else 
		{
			cell* decoy=dl->left;
			dl->left = decoy->left;
			decoy->left->right = dl;
			delete[]decoy;
		}

	}

	void deleting_column_left()
	{
		bool fl=false;
		XL_Iterator decoy(this->current_cell); // saving the cell at which we are right now
		if (this->current_cell.current->left != nullptr) // because agr left pa koi cell nhi to koi column delete nhi ho ga or grid chota nhi ho ga
		{
			this->current_col--;
			this->grid_col_dim--;
		}

		// going at the start of the column in which our current cell is present
		while (decoy.current->up != nullptr)
		{
			--decoy; // going up
		}
		if (*(decoy.current->left) == *this->head_pointer)
		{   // if the column about to be deleted contains the head pointer shift it to its right cell
			this->head_pointer == decoy.current;
		}

		XL_Iterator top(decoy);// top cell of the column ko save kar rha hon

		while (decoy.current != nullptr)
		{
			this->deleting_left_cell(decoy.current);
			++decoy; // going down
		}
	}

	//______   deleting rows
	// upward
	void deleting_cell_above(cell* dl)
	{
		if (dl->up == nullptr)
		{
			return;
		}
		else if (dl->up->up == nullptr)
		{
			cell* decoy = dl->up;
			dl->up = nullptr;
			delete[]decoy;
			
		}
		else 
		{
			cell* decoy = dl->up;
			decoy->up->down = dl;
			dl->up = decoy->up;
			delete[]decoy;
		}
	}

	void delete_row_above()
	{
		bool fl = false;
		XL_Iterator decoy(this->current_cell); // saving the cell at which we are right now
		if (this->current_cell.current->up != nullptr) // because agr iper koi cell nhi to koi row delete nhi ho ga or grid chota nhi ho ga
		{
			// going at the start of the column in which our current cell is present
			while (decoy.current->left != nullptr)
			{
				decoy--; // going left
			}
			if (*(decoy.current->up) == *this->head_pointer)
			{   // if the column about to be deleted contains the head pointer shift it to its right cell
				this->head_pointer == decoy.current;
			}

			XL_Iterator left_most(decoy);// top cell of the column ko save kar rha hon

			while (decoy.current != nullptr)
			{
				this->deleting_cell_above(decoy.current);
				decoy++; // going right
			}
			this->current_row--;
			this->grid_row_dim--;
		}
	}

	// downward

	void delete_cell_below(cell* dl)
	{
		if (dl->down == nullptr)
		{
			return;
		}
		else if (dl->down->down == nullptr)
		{
			cell* decoy = dl->down;
			dl->down = nullptr;
			delete[]decoy;
			
		}
		else
		{
			cell* decoy = dl->down;
			decoy->down->up = dl;
			dl->down = decoy->down;
			delete[]decoy;
		}
	}

	void delete_row_below()
	{
		XL_Iterator decoy(this->current_cell); // saving the cell at which we are right now
		if (this->current_cell.current->down != nullptr) // because agr down pa koi cell nhi to koi column delete nhi ho ga or grid chota nhi ho ga
		{
			// going at the start of the column in which our current cell is present
			while (decoy.current->left != nullptr)
			{
				decoy--; // going left
			}
			XL_Iterator left_most(decoy);// top cell of the column ko save kar rha hon

			while (decoy.current != nullptr)
			{
				this->delete_cell_below(decoy.current);
				decoy++; // going right
			}
			this->grid_row_dim--;
		}

	}

	void clear_whole_row()
	{
		XL_Iterator decoy(this->current_cell);
		// going at the start of row;
		while (decoy.current->left != nullptr)
		{
			decoy--;  // going left
		}
		// now going right and also clearing the whole row
		while (decoy.current != nullptr)
		{
			decoy.current->data.clear();
			decoy++; // going right
		}
	}

	void clear_whole_column()
	{
		XL_Iterator decoy(this->current_cell);
		// going at the start of row;
		while (decoy.current->up != nullptr)
		{
			--decoy;  // going up
		}
		// now going right and also clearing the whole row
		while (decoy.current != nullptr)
		{
			decoy.current->data.clear();
			++decoy; // going down
		}
	}



	//  singl cell insertion and deletion

	//__________   Inserting a single cell

	void insert_single_cell_by_right_shift()
	{
		// saving current cell
		XL_Iterator decoy_current(this->current_cell);
		while (this->current_cell.current->right != nullptr)
		{
			this->current_cell++;  // going right
		}
		// now using already existing column insert function
		this->InsertColumnAtRight(); // this uses current iterator 
		
		// making an iterator pointing at the current cell which is actually the cell before the last cell of the row 
		XL_Iterator right_most(this->current_cell.current);
		this->current_cell.set_pointer(decoy_current.current); // re position the current pointer to its original position

		// now shifting the data of cells into the cell 
		while (right_most.current != this->current_cell.current)
		{
			right_most.current->right->data = right_most.current->data; // shifting to right
			right_most--; // going left
		}
		right_most.current->right->data.clear();
		
	}

	void insert_single_cell_by_left_shift() 
	{
		// saving current cell
		XL_Iterator decoy_current(this->current_cell);
		while (this->current_cell.current->left != nullptr)
		{
			this->current_cell--;  // going left
		}
		// now using already existing column insert function
		this->InsertColumnAtLeft(); // this uses current iterator 

		// making an iterator pointing at the current cell which is actually the cell before the last cell of the row 
		XL_Iterator left_most(this->current_cell.current);
		left_most--; 
		this->current_cell.set_pointer(decoy_current.current); // re position the current pointer to its original position
		//this->current_col++;  
		// now shifting the data of cells into the cell 
		while (left_most.current->right != this->current_cell.current)
		{
			left_most.current->data = left_most.current->right->data; // shifting to right
			left_most++; // going right
		}
		left_most.current->data.clear();
	}

	void insert_single_cell_by_down_shift()
	{
		XL_Iterator decoy_current(this->current_cell);
		while (this->current_cell.current->down != nullptr)
		{
			++this->current_cell;  // going down
		}
		// now using already existing column insert function
		this->InsertRowBelow(); // this uses current iterator 

		// making an iterator pointing at the current cell which is actually the cell before the last cell of the row 
		XL_Iterator down_most(this->current_cell.current);
		++down_most;
		this->current_cell.set_pointer(decoy_current.current); // re position the current pointer to its original position
		//this->current_col++;  
		// now shifting the data of cells into the cell 
		while (down_most.current->up != this->current_cell.current)
		{
			down_most.current->data = down_most.current->up->data; // shifting to right
			--down_most; // going up
		}
		down_most.current->data.clear();
	}

	void insert_single_cell_by_up_shift()
	{
		XL_Iterator decoy_current(this->current_cell);
		while (this->current_cell.current->up != nullptr)
		{
			--this->current_cell;  // going up
		}
		// now using already existing column insert function
		this->InsertRowAbove(); // this uses current iterator 

		// making an iterator pointing at the current cell which is actually the cell before the last cell of the row 
		XL_Iterator up_most(this->current_cell.current);
		--up_most;
		this->current_cell.set_pointer(decoy_current.current); // re position the current pointer to its original position
		//this->current_col++;  
		
		// now shifting the data of cells into the cell 
		while (up_most.current->down != this->current_cell.current)
		{
			up_most.current->data = up_most.current->down->data; // shifting to right
			++up_most; // going down
		}
		up_most.current->data.clear();
	}

	// deletion of a single cell by shifting
	 
	void delete_cell_at_right()
	{
		if (this->current_cell.current->right == nullptr) // agr right pa cell hi nhi too delete kia karna 
		{
			return;
		}
		else if (this->current_cell.current->right->right == nullptr)
		{ 
			this->current_cell.current->right->data.clear();  // right wla cell empty kar do
		}
		else
		{
			XL_Iterator decoy(this->current_cell);
			decoy++; // 1 step right
			while (decoy.current->right != nullptr)
			{
				decoy.current->data = decoy.current->right->data;
				decoy++;  // going right
			}
			decoy.current->data.clear();
		}
	}

	void delete_cell_at_left()
	{
		if (this->current_cell.current->left == nullptr) // agr right pa cell hi nhi too delete kia karna 
		{
			return;
		}
		else if (this->current_cell.current->left->left == nullptr)
		{
			this->current_cell.current->left->data.clear();  // right wla cell empty kar do
		}
		else
		{
			XL_Iterator decoy(this->current_cell);
			decoy--; // 1 step left
			while (decoy.current->left!=nullptr)
			{
				decoy.current->data = decoy.current->left->data;
				decoy--; // going left
			}
			decoy.current->data.clear();
		}
	}

	void delete_cell_above()
	{
		if (this->current_cell.current->up == nullptr) // agr right pa cell hi nhi too delete kia karna 
		{
			return;
		}
		else if (this->current_cell.current->up->up == nullptr)
		{
			this->current_cell.current->up->data.clear();  // right wla cell empty kar do
		}
		else
		{
			XL_Iterator decoy(this->current_cell);
			--decoy; // 1 step up
			while (decoy.current->up != nullptr)
			{
				decoy.current->data = decoy.current->up->data;
				--decoy;
			}
			decoy.current->data.clear();
		}
	}

	void delete_cell_below()
	{
		if (this->current_cell.current->down == nullptr) // agr right pa cell hi nhi too delete kia karna 
		{
			return;
		}
		else if (this->current_cell.current->down->down == nullptr)
		{
			this->current_cell.current->down->data.clear();  // right wla cell empty kar do
		}
		else
		{
			XL_Iterator decoy(this->current_cell);
			++decoy; // 1 step up
			while (decoy.current->down != nullptr)
			{
				decoy.current->data = decoy.current->down->data;
				++decoy;  // going down
			}
			decoy.current->data.clear();
		}
	}

	//____________    function   for  selecting range

	



public:
	// Default Constructor
	Mini_Excel()
		:head(), current_cell()
	{	
		this->head_pointer = &this->head;
		this->current_row = 0;
		this->current_col = 0;
		this->is_a_cell_selected = false;
		cout << "\n\n\n        Do you want to load the previous sheet (Y/N): ";
		char input;
		cin >> input;
		system("cls");	
		this->current_cell.set_pointer(this->head_pointer);// 1st cell is the attribute of Mini_Excel class because we need a starting
		if (input == 'Y' || input == 'y')
		{
			int row_to_reach = 0, col_to_reach = 0;
			this->grid_row_dim = 1;
			this->grid_col_dim = 1;
			fstream rdr("bscs22025_exceL_save_sheet.txt");
			rdr >> row_to_reach;
			rdr >> col_to_reach;
			this->grid_row_dim = 1;
			this->grid_col_dim = 1;
			int getline_index = 0;
			XL_Iterator decoy_it(current_cell);

			for (int i = 1; i < col_to_reach; i++)				// we are going to initialize 4 cols
			{										// since 1 head cell is already present
				this->InsertColumnAtRight();
				current_cell++; // moving right
			}
			for (int j = 1; j < row_to_reach; j++)  // A row is already formed in above loop as the 1st row
			{
				this->InsertRowBelow();
				++current_cell;// moving down
			}
			current_cell.set_pointer(decoy_it.current);
			XL_Iterator going_down(this->current_cell);
			string s;
			getline(rdr, s);
			for (int i = 0; i < row_to_reach; i++)
			{
				getline(rdr, s);
				XL_Iterator going_right(going_down);
				for (int j = 0; j < col_to_reach; j++)
				{
					string inp="";
					while(true)
					{
						if (s[getline_index] == ' ' || s[getline_index] == '\n')
						{
							getline_index++;
							break;
						}
						inp.push_back(s[getline_index]);
						getline_index++;
					}
					if (inp != "_")
					{
						going_right.current->data = inp;;
					}
					going_right++;// going right
				}
				getline_index = 0;
				++going_down; // going down
			}
		}
		else
		{
			
			// cell
			this->grid_row_dim = 1;
			this->grid_col_dim = 1;
			XL_Iterator decoy_it(current_cell);

			for (int i = 1; i < 5; i++)				// we are going to initialize 4 cols
			{										// since 1 head cell is already present
				this->InsertColumnAtRight();
				current_cell++; // moving right
			}
			for (int j = 1; j < 5; j++)  // A row is already formed in above loop as the 1st row
			{
				this->InsertRowBelow();
				++current_cell;// moving down
			}
			current_cell.set_pointer(decoy_it.current); // resetting the pointer to the start

		}
	}

	//+++++++++++++++++++++    deletion functions of current cols and rows  +++++++++++++++++++++++++++
	void delete_current_row()
	{
		XL_Iterator decoy(this->current_cell.current);
		while (decoy.current->left != nullptr)
		{
			decoy--; // going left
		}
		if (decoy.current == this->head_pointer)  // agr jo row delete karna laha han wo head pa ha too head ko bhi shift kar doo
		{
			this->head_pointer = this->head_pointer->down; // shifting to down cell
		}
		// shifting the current cell
		if (this->current_cell.current->down == nullptr)
		{
			this->current_cell.current = this->current_cell.current->up;
			this->current_row--;
		}
		else
		{
			this->current_cell.current = this->current_cell.current->down;
			
		}
		//  Now deleting the whole row
		while (decoy.current != nullptr)
		{
			if (decoy.current->up!=nullptr)
			{
				decoy.current->up->down = decoy.current->down;
			}
			if (decoy.current->down !=nullptr)
			{
				decoy.current->down->up = decoy.current->up;
			}
			decoy.current->up = nullptr; decoy.current->down = nullptr;
			// if we are in the middle and we have to 1st shift the decoy to its right
			if (decoy.current->right != nullptr)
			{
				cell* del = decoy.current;
				decoy++;
				decoy.current->left = nullptr;
				del->left = nullptr; del->right = nullptr;
   				//delete[]del;
			} // if we are already at the end then we are not going to shift and only delete decoy
			else
			{    
				decoy.current->left = nullptr;
				//delete[]decoy.current;
				break;
			}
			
		}
		this->grid_row_dim--;
	}

	void delete_current_column()
	{
		XL_Iterator decoy(this->current_cell.current);
		while (decoy.current->up != nullptr)
		{
			--decoy; // going up
		}
		if (decoy.current == this->head_pointer)  // agr jo row delete karna laha han wo head pa ha too head ko bhi shift kar doo
		{
			this->head_pointer = this->head_pointer->right; // shifting to right cell
		}
		// shifting the current cell
		if (this->current_cell.current->right == nullptr)
		{
			this->current_cell.current = this->current_cell.current->left;
			this->current_col--;
		}
		else
		{
			this->current_cell.current = this->current_cell.current->right;
		}
		//  Now deleting the whole column
		while (decoy.current != nullptr)
		{
			if (decoy.current->left != nullptr)
			{
				decoy.current->left->right = decoy.current->right;
			}
			if (decoy.current->right != nullptr)
			{
				decoy.current->right->left = decoy.current->left;
			}
			decoy.current->left = nullptr; decoy.current->right = nullptr;
			// if we are in the middle and we have to 1st shift the decoy to its down
			if (decoy.current->down != nullptr)
			{
				cell* del = decoy.current;
				++decoy; // going down
				decoy.current->up = nullptr;
				del->down = nullptr; del->up = nullptr;
				//delete[]del;
			}
			// if we are already at the end then we are not going to shift and only delete decoy
			else
			{
				decoy.current->up = nullptr;
				//delete[]decoy.current;
				break;
			}

		}
		this->grid_col_dim--;
	}
	
	bool is_alphabet(char x)
	{
		if (x == '.' || (x >= '1' && x <= '9')) // its not an alphabet
		{
			return false;
		}
		else
			return true; // its  an alphabet
	}

	void value_inserter_for_range_function(cell* x,vector<float>& values)
	{
		string* val_to_check = &x->data;
		if (!val_to_check->empty())
		{
			if (!is_alphabet((*val_to_check)[0]))
			{
				values.push_back(stof(x->data));  // agr alpha bet nhi ha to number ho ga
			}// number ka alava kuch puch karwana ki zarorat nhi
		}
	}

	vector<float> range_values_collector(int& range_start_row, int& range_start_col, int& range_end_col, int& range_end_row, char keyboard_command)   // this function is for setting the start and end of the range 
	{
		vector<float> values;
		// if single cell is selected
		string* val_to_check = &this->range_start.current->data;
		if (range_start_row == range_end_row && range_end_col==range_start_col)
		{    // agr cell ka data empty nhi ha too 
			value_inserter_for_range_function(this->range_start.current,values);
			return values;
		}  // agr row ak hi ho par column different hon (horizontal line case)
		else if (range_end_row == range_start_row)
		{    // agr starting col, ending col (coordinates)  sa bara ha to start or end ko swap kar da
			if (range_start_col > range_end_col)
			{
				// because cell contains pointers will this work ?
				swap(this->range_start.current, this->range_end.current);
				swap(range_start_col, range_end_col);
			}
			XL_Iterator going_right(this->range_start.current);
			while (range_start.current != range_end.current)
			{
				value_inserter_for_range_function(going_right.current, values);
				range_start++;
				going_right++; // going right
			}
			value_inserter_for_range_function(going_right.current, values);
			return values;
		}// same column but change rows
		else if (range_end_col == range_start_col)
		{
			if (range_start_row > range_end_row)
			{
				// because cell contains pointers will this work ?
				swap(this->range_start.current, this->range_end.current);
				swap(range_start_row, range_end_row);
			}
			XL_Iterator going_down(this->range_start.current);
			while (range_start.current != range_end.current)
			{
				value_inserter_for_range_function(going_down.current, values);
				++range_start;
				++going_down;
			}
			value_inserter_for_range_function(going_down.current, values);
			return values;
		}  // both rows and columns are different from each other
		else 
		{   // agr range ka start box range ka end box sa aga ha
			if (range_end_col < range_start_col && range_end_row < range_start_col)
			{
				swap(this->range_start.current, this->range_end.current);
				swap(range_start_row, range_end_row);
				swap(range_start_col, range_end_col);
			}   // going down in rows
			if (range_start_col<range_end_col && range_start_row>range_end_row)
			{
				for (int i = range_start_row; i > range_end_row; i--)
				{
					this->range_start--; // going up
					this->range_end++;  // going down
				}
				swap(range_start_row, range_end_row);
			}
			XL_Iterator decoy_start(this->range_start);
			for (int i = 0; i <= range_end_row-range_start_row; i++)
			{    // going from left to right
				XL_Iterator traverser(decoy_start);
				for (int j = 0; j <=range_end_col-range_start_col; j++)
				{
					value_inserter_for_range_function(traverser.current, values);
					traverser++; // going right
				}
				++decoy_start; // going down after storing the values of row infront
			}
		}
		return values;
	}


	bool Range_selector(int &range_start_row,int& range_start_col , int& range_end_col, int& range_end_row, char keyboard_command)
	{
		range_start_row = this->current_row, range_start_col = this->current_col;
		this->range_start.set_pointer(this->current_cell.current);
		while (true)  // end the range
		{
			keyboard_command = _getch();
			// 1st we are going to select a range
			if (int(keyboard_command) == -32)
			{
				keyboard_command = _getch();
				current_cell_mover(keyboard_command);
				this->current_cell_printer(this->current_cell.current, this->current_row, this->current_col);
			}
			else if (keyboard_command == 'E' || keyboard_command == 'e')  // end range function and perform the function on the selected range
			{
				range_end_row = this->current_row; range_end_col = this->current_col;
				this->range_end.set_pointer(this->current_cell.current);
				return true;// continue the range function
			}
			else  // terminate the range operations
			{
				// reset range and come out
				this->range_start.set_pointer(nullptr);
				range_start_row = 0, range_end_col = 0;
				return false;  // terminate range functions
			}
		}
	}

	void copy_cut(int range_start_row, int range_start_col, int range_end_row, int range_end_col, bool is_cut = false)  // using for copying and cutting data in cells
	{
		char cmd;

		//+++++++++++++++++++++++++++++++++++++++++++++++++
		 // for same row (horizontal case only)
		//+++++++++++++++++++++++++++++++++++++++++++++++++

		if (range_start_row == range_end_row)
		{   // agr range right to left select ki ha to unha invert kar da
			if (range_start_col > range_end_col)
			{
				//this->range_start.current->swap_cells(*this->range_end.current);
				swap(this->range_start.current, this->range_end.current);
				swap(range_start_col, range_end_col);
			}
			vector<string> horizontal_copy;
			XL_Iterator going_right(this->range_start);
			XL_Iterator cut_data(this->range_start);
			// copying 
			while (going_right.current != this->range_end.current)
			{
				horizontal_copy.push_back(going_right.current->data);
				going_right++; // going right
			}
			horizontal_copy.push_back(going_right.current->data);
			// now we are going to choose the cell from which we are going to paste the data onward
			int decoy_current_col = this->current_col;
			while (true)
			{
				cmd = _getch();
				
				if (int(cmd) == -32)
				{
					cmd = _getch();
					current_cell_mover(cmd);
					this->current_cell_printer(this->current_cell.current, this->current_row, this->current_col);
				}
				else if (int(cmd) == 22)  // ctrl v
  				{
					XL_Iterator decoy_current(this->current_cell);
					if (is_cut)  // agr cut option choose ki ha to wo data copy karna ka sath delete bhi kar da ga
					{
						while (cut_data.current != this->range_end.current)
						{
							cut_data.current->data.clear();
							cut_data++;
						}
						cut_data.current->data.clear();
					}
					// pasting data horizontally
					for (int i = 0; i < horizontal_copy.size(); i++)
					{
						
						if (decoy_current_col <= 10)
						{
							this->current_cell.current->data = horizontal_copy[i];
							// agr ab right pa koi cell nhi too new column dalo
							if (this->grid_col_dim<=10 &&this->current_cell.current->right == nullptr && i < horizontal_copy.size() - 1)
							{
								this->InsertColumnAtRight();
							}
							this->current_cell++; // going right
							decoy_current_col++;
						}
					}
					this->current_cell.set_pointer(decoy_current.current);
					return;
				}
				else
				{
					return;  // just terminate the copy paste function
				}
			}
		}
		
		//+++++++++++++++++++++++++++++++++++++++++++++++++
		// for verticle case  (same column)
		//+++++++++++++++++++++++++++++++++++++++++++++++++

		else if (range_start_col == range_end_col)
		{   // agr range down to up select ki ha to unha invert kar da
			if (range_start_row > range_end_row)
			{
				//this->range_start.current->swap_cells(*this->range_end.current);
				swap(this->range_start.current, this->range_end.current);
				swap(range_start_row, range_end_row);
			}
			vector<string> verticle_copy;
			XL_Iterator going_down(this->range_start);
			XL_Iterator cut_data(this->range_start); // if cutting is enabled
			// copying 
			while (going_down.current != this->range_end.current)
			{
				verticle_copy.push_back(going_down.current->data);
				++going_down; // going down
			}
			verticle_copy.push_back(going_down.current->data);
			// now we are going to choose the cell from which we are going to paste the data onward
			while (true)
			{
				cmd = _getch();
				if (int(cmd) == -32)
				{
					cmd = _getch();
					current_cell_mover(cmd);
					this->current_cell_printer(this->current_cell.current, this->current_row, this->current_col);
				}
				else if (int(cmd) == 22 )  // ctrl v
				{
					XL_Iterator decoy_current(this->current_cell);
					int decoy_curr_row = this->current_row;
					if (is_cut)  // agr cut option choose ki ha to wo data copy karna ka sath delete bhi kar da ga
					{
						while (cut_data.current != this->range_end.current)
						{
							cut_data.current->data.clear();
							++cut_data; // going down
						}
						cut_data.current->data.clear();
					}
					// pasting data horizontally
					for (int i = 0; i < verticle_copy.size(); i++)
					{
						if (decoy_curr_row<=19)
						{
							this->current_cell.current->data = verticle_copy[i];
							// agr ab right pa koi cell nhi too new column dalo
							if (this->grid_row_dim<=19 && this->current_cell.current->down == nullptr && i < verticle_copy.size() - 1)
							{
								this->InsertRowBelow();
							}
							++this->current_cell; // going down
							decoy_curr_row++;
						}
						
					}
					this->current_cell.set_pointer(decoy_current.current);
					return;
				}
				else
				{
					return;  // just terminate the copy paste function
				}
			}
		}
	}

	float range_sum(vector<float>values)
	{
		float sum = 0;
		for (auto i : values)
		{
			sum += i;
		}
		return sum;
	}

	float range_average(const vector<float> &values)
	{
		float avg = range_sum(values);
		avg /= values.size();
		return avg;
	}

	float count(const vector<float>& values)
	{
		return values.size();
	}

	float Min(const vector<float>& values)
	{
		auto i= min_element(values.begin(), values.end());
		return float(*i);
	}
	
	float Max(const vector<float>& values)
	{
		auto i = max_element(values.begin(), values.end());
		return float(*i);
	}

	// saving data on a new file
	void save_file()
	{
		fstream wrt("bscs22025_exceL_save_sheet.txt", std::ofstream::out | std::ofstream::trunc);
		wrt << this->grid_row_dim << " " << this->grid_col_dim << endl;
		XL_Iterator going_down(this->head_pointer);
		for (int i = 0; i < this->grid_row_dim; i++)
		{
			XL_Iterator going_right(going_down);
			for (int j = 0; j < this->grid_col_dim; j++)
			{
				if (going_right.current->data == "")
				{
					wrt << "_" << " ";  // if the cell was empty
				}
				else
				{
					wrt << going_right.current->data << " ";
				}
				going_right++;
			}
			wrt << endl;
			++going_down;
		}
	}



	//++++++++++  the main running function of excel

	// Remember :
	/*  The maximum number of cell columns can be 11 and the maximum
		number of rows of cells can be the height of the console*/
	
	void use_Excel()
	{
		
		char keyboard_command;
		int range_start_col=0, range_start_row=0;
		int range_end_col = 0, range_end_row = 0;

		this->whole_sheet_printer(this->head_pointer,this->grid_row_dim,this->grid_col_dim);
		
		while (true) // jab ma close nhi karta window jab tak chalta rha 
		{
			this->current_cell_printer(this->current_cell.current, this->current_row, this->current_col);
			keyboard_command = _getch(); // taking input from keyboard

			// when an arrow key is pressed and you want to move the current cell in its direction
			if (int(keyboard_command) == -32)
			{
				keyboard_command = _getch();
				current_cell_mover(keyboard_command);

			}// if the user has selected a cell to type in it
			else if (keyboard_command == '/')
			{
				this->is_a_cell_selected = true;
				while (is_a_cell_selected)
				{
					keyboard_command = _getch();  // taking input for what is going to be written in the cell
					if (keyboard_command == '/')
					{
						this->is_a_cell_selected = false;
						break; // this means come out of the cell selected mode
					}
					else  // somthing is about to enter the cell data
					{     // if the input is valid
						if (is_input_data_valid(this->current_cell.current, keyboard_command))
						{
							enter_and_display_data_in_cell(this->current_cell.current, this->current_row, this->current_col, keyboard_command);
						}
					}

				}

			}  // agr column ya row add karna ka kaha ha add 
			else if (keyboard_command == 'I' || keyboard_command == 'i')  // i stands for insert
			{
				keyboard_command = _getch();
				if (keyboard_command == 'R' || keyboard_command == 'r') // r stands for row
				{
					keyboard_command = _getch();
					if (this->grid_row_dim <= 19)  // max rows that can be inserted are 19
					{
						if (keyboard_command == 'A' || keyboard_command == 'a') // a stands for above
						{

							this->InsertRowAbove();
							//system("cls");
							this->whole_sheet_printer(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
						}
						else if (keyboard_command == 'B' || keyboard_command == 'b') // b stands for below
						{
							this->InsertRowBelow();
							//system("cls");
							this->whole_sheet_printer(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
						}
					}

				}
				else if (keyboard_command == 'C' || keyboard_command == 'c')  // c stands for column
				{
					keyboard_command = _getch();
					if (this->grid_col_dim <= 10)
					{
						if (keyboard_command == 'R' || keyboard_command == 'r') // r stands for right
						{
							this->InsertColumnAtRight();
							//system("cls");
							this->whole_sheet_printer(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
						}
						else if (keyboard_command == 'L' || keyboard_command == 'l') // l stands for left
						{
							this->InsertColumnAtLeft();
							//system("cls");
							this->whole_sheet_printer(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
						}
					}
				}
			} // if we want to delete  a column or a row
			else if (keyboard_command == 'D' || keyboard_command == 'd')
			{
				keyboard_command = _getch();
				if (keyboard_command == 'R' || keyboard_command == 'r') // r stands for row
				{
					
					if (this->grid_row_dim > 3)  
					{
						this->delete_current_row();
						system("cls");
						this->whole_sheet_printer(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
					}

				}
				else if (keyboard_command == 'C' || keyboard_command == 'c')  // c stands for column
				{
					
					if (this->grid_col_dim > 3)
					{
						this->delete_current_column();
						system("cls");
						this->whole_sheet_printer(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
					}
				}
			}
			else if (keyboard_command == 'c' || keyboard_command == 'C') // cell change mode  (like string alligment, color)
			{
				int pr = 100, pc = 105;
				keyboard_command = _getch();
				if (keyboard_command == 'k' || keyboard_command == 'K')  // k is for color change
				{
					

					// 100 ,122
					gotoRowCol(pr, pc);
					char klr;
					SetClr(15);
					cout << "Enter the colour of the cell text you want, options:(White:w , green:g , blue:b) : ";
					klr=_getch();;
					while (!(klr == 'w' || klr == 'W' || klr == 'g' || klr == 'G' || klr == 'b' || klr == 'B'))
					{
						gotoRowCol(pr, pc);
						cout << "                                                                               ";
						gotoRowCol(pr, pc);
						cout << "You have entered invalid colour please enter again: ";
						klr=_getch();
					}
					gotoRowCol(pr, pc);
					cout << "                                                                                   ";
					if (klr == 'w' || klr == 'W')
					{
						this->current_cell.current->text_clr = WHITE;
					}
					else if (klr == 'g' || klr == 'G')
					{
						this->current_cell.current->text_clr = GREEN;
					}
					else if (klr == 'b' || klr == 'B')
					{
						this->current_cell.current->text_clr = LBLUE;
					}
				}
				else if (keyboard_command == 's' || keyboard_command == 'S') // s for style change of the text
				{
					gotoRowCol(pr, pc);
					char klr;
					SetClr(15);
					cout << "Enter the Alignment type of text you want (c: centre, r: right, l:left): ";
					klr=_getch();;
					while (!(klr == 'R' || klr == 'r' || klr == 'c' || klr == 'C' || klr == 'L' || klr == 'l'))
					{
						gotoRowCol(pr, pc);
						cout << "                                                                                          ";
						gotoRowCol(100, 115);
						cout << "You have entered invalid alignnmet type please enter again: ";
						klr=_getch();;
					}
					gotoRowCol(pr, pc);
					cout << "                                                                                                 ";
					if (klr == 'r' || klr == 'R')
					{
						this->current_cell.current->cell_alignment = 'r';
					}
					else if (klr == 'c' || klr == 'C')
					{
						this->current_cell.current->cell_alignment = 'c';
					}
					else if (klr == 'L' || klr == 'l')
					{
						this->current_cell.current->cell_alignment = 'l';
					} 
				}


			}   
			
			//++++++++++++++  E stands for erase command (this will be used to erase the data of a column or row) +++++++++++++
			
			else if (keyboard_command == 'e' || keyboard_command == 'E')
			{
				keyboard_command = _getch();
				if (keyboard_command == 'c' || keyboard_command == 'C')
				{
					clear_whole_column();
					this->whole_sheet_printer(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
				}
				else if (keyboard_command == 'r' || keyboard_command == 'R')
				{
					clear_whole_row();
					this->whole_sheet_printer(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
				}
			}
			
			//++++++++++++++++++++++++++++++++++   insert 1 (cell)    ++++++++++++++++++++++++++++++++++++++++++++++++++=
			
			else if (keyboard_command == '!' || keyboard_command == '1')
			{
				keyboard_command = _getch();
				// inserting in column side
				if (keyboard_command == 'i' || keyboard_command == 'I')
				{
					keyboard_command = _getch();
					if (this->grid_col_dim <= 10)
					{
						if (keyboard_command == 'r' || keyboard_command == 'R')  // 1 cell at right with right shift
						{
							this->insert_single_cell_by_right_shift();
							this->whole_sheet_printer(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
						}
						else if (keyboard_command == 'l' || keyboard_command == 'L')
						{
							this->insert_single_cell_by_left_shift();
							this->whole_sheet_printer(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
						}
					}
					if (this->grid_row_dim <= 19)
					{
						if (keyboard_command == 'a' || keyboard_command == 'A')  // 1 cell up with up shift
						{
							this->insert_single_cell_by_up_shift();
							this->whole_sheet_printer(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
						}
						else if (keyboard_command == 'b' || keyboard_command == 'B') // 1 cell down with down shift
						{
							this->insert_single_cell_by_down_shift();
							this->whole_sheet_printer(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
						}
					}
				}  // deleting a single cell
				else if (keyboard_command == 'd' || keyboard_command == 'D')
				{
					//++++++++++=    insert and check the single cell deletion functions
					keyboard_command = _getch();
					
					
					if (keyboard_command == 'r' || keyboard_command == 'R')  // 1 cell at right with right shift
					{
						this->delete_cell_at_right();
						this->whole_sheet_printer(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
					}
					else if (keyboard_command == 'l' || keyboard_command == 'L')
					{
						this->delete_cell_at_left();
						this->whole_sheet_printer(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
					}
					
					if (keyboard_command == 'A' || keyboard_command == 'a')  // 1 cell up with up shift
					{
						this->delete_cell_above();
						this->whole_sheet_printer(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
					}
					else if (keyboard_command == 'B' || keyboard_command == 'b') // 1 cell down with down shift
					{
						this->delete_cell_below();
						this->whole_sheet_printer(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
					}
				}
				
				
			}

			// This is the mouse mode
			else if (keyboard_command == 'm' || keyboard_command == 'M')
			{	// agr koi range select hoi ha too proceed karo warna terminate karka bahir aa jao
				
				if (Range_selector(range_start_row, range_start_col, range_end_col, range_end_row, keyboard_command)) 
				{
					vector<float> value=range_values_collector(range_start_row,range_start_col,range_end_col,range_end_row,keyboard_command);
					int ans = 0;
					bool op_done = false;
					int row, col;
					getRowColbyLeftClick(row, col);
					if (col >= 13 * cell_col_dim && col <= 14 * cell_col_dim)
					{
						if (row>=2*cell_row_dim && row<=3*cell_row_dim)
						{
							ans = this->range_sum(value);
							op_done = true;
						}
						else if (row >= 4 * cell_row_dim && row <= 5 * cell_row_dim)
						{
							ans=this->count(value);
							op_done = true;
						}
						else if (row >= 6 * cell_row_dim && row <= 7 * cell_row_dim)
						{
							ans=this->range_average(value);
							op_done = true;
						}
						else if (row >= 8 * cell_row_dim && row <= 9 * cell_row_dim)
						{
							ans=this->Min(value);
							op_done = true;
						}
						else if (row >= 10* cell_row_dim && row <= 11 * cell_row_dim)
						{
							ans = this->Max(value);
							op_done = true;
						}
						
						if (op_done)  // op_done means operation done
						{
							while (true)  
							{
								keyboard_command = _getch();
								if (keyboard_command == '/')
								{
									this->current_cell.current->data.clear();
									this->current_cell.current->data =to_string(ans);
									break;
								}
								current_cell_mover(keyboard_command);
								this->current_cell_printer(this->current_cell.current, this->current_row, this->current_col);

							}
						}
						
						// discard 	
						value.clear();

					}
				}
			}
			// ++++++++++++     copy/ cut paste
			else if (int(keyboard_command) == 3 || int(keyboard_command) == 24) //  ctrl + (c=3  and   x=4)
			{    // if range is selected then 
				if (Range_selector(range_start_row, range_start_col, range_end_col, range_end_row, keyboard_command))
				{
					if (keyboard_command==3) // copy 
					{
						copy_cut(range_start_row, range_start_col, range_end_row, range_end_col);
						this->whole_sheet_printer(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
					}
					else // cut
					{
						copy_cut(range_start_row, range_start_col, range_end_row, range_end_col, true);
						this->whole_sheet_printer(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
					}
				}
			}
			//WriteAllText("")
			this->save_file();
		}
		
	}

};










