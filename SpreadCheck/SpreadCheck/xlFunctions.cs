using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SpreadCheck
{
    class XLFunctions
    {
        public string TrimCell(string Cell)
        {   if (!String.IsNullOrEmpty(Cell))
                return Cell.Trim(); 
            else return Cell;        
        }

        public string ToUpperCase(string Cell)
        {  if (!String.IsNullOrEmpty(Cell))
                return Cell.ToUpper();
            else return Cell;
        }

        public string ChangeCase(string Cell, int ChangeType)
        {	if (!String.IsNullOrEmpty(Cell))
                {   switch (ChangeType)
                    {
					case 0:
                        return Cell;
					break;
                    case 1:
                        return Cell.ToUpper();
                        break;
                    case 2:
                        return Cell.ToLower();
                        break;
                    case 3:
                        return Cell.ToUpperInvariant();
                        break;
                    case 4:
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
        {   if (!String.IsNullOrEmpty(Cell))
                return Cell.ToLower();
            else return Cell;
        }

        public string ReverseCell(string Cell)
        {  if (!String.IsNullOrEmpty(Cell))
                return new string(Cell.Reverse().ToArray());
            else return Cell;
        }
    }
}
