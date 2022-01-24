using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SpreadCheck
{
	public class XLFunctions
	{
		public int FunctionCallCount;

		public object TrimCell(object Cell)
		{	switch (Cell) {
				case string s:

				if (!String.IsNullOrEmpty(s)) {
					FunctionCallCount++;
					return s.Trim();
				}
				else return s;
				break;

				default:
				break;
			}
			return Cell;

        }


		public object ToUpperCase(object Cell)
		{
			switch (Cell) {

				case string s:

				if (!String.IsNullOrEmpty(s)) {
					FunctionCallCount++;
					return s.ToUpper();
				}

				else return Cell;
				break;

				default:
				break;
			}

			return Cell;
        }

        public string ChangeCase(string Cell, int ChangeType)
        {	if (!String.IsNullOrEmpty(Cell))
                {   switch (ChangeType)
                    {
					case 0:
						FunctionCallCount++;
						return Cell;
						break;
                    case 1:
						FunctionCallCount++;
						return Cell.ToUpper();
                        break;
                    case 2:
						FunctionCallCount++;
						return Cell.ToLower();
                        break;
                    case 3:
						FunctionCallCount++;
						return Cell.ToUpperInvariant();
                        break;
                    case 4:
						FunctionCallCount++;
						return Cell.ToLowerInvariant();
                        break;
                    default:
                        return Cell;
                        break;
                    }
            }
            return null;
        }

        public string ToLowerCase(string Cell)
        {   if (!String.IsNullOrEmpty(Cell)) {
				FunctionCallCount++;
				return Cell.ToLower();
			}
			else return Cell;
        }

        public string ReverseCell(string Cell)
        {  if (!String.IsNullOrEmpty(Cell)) {
				FunctionCallCount++;
				return new string(Cell.Reverse().ToArray());
			}
			else return Cell;
        }
    }
}
