using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ComInherite
{
	/*----------------------------------------------------------------------------------------
     * U kazdeho interfacu a tridy musi byt vygenerovany GUID a [ComVisible(true)]
     * u trid [ClassInterface(ClassInterfaceType.None)] a dedit i z MarshalByRefObject
     * u interface kde ma trida neco vracet z delphi (prepise se ji v delphi metoda) [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    ---------------------------------------------------------------------------------------*/

	[Guid("898B0A63-BC76-43C6-BED0-21F485367701")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	[ComVisible(true)]
	public interface ICalc
	{
		IDelphiEventContainer DelphiEvents { get; set; }
		int Sum(int a, int b);
	}

	[Guid("F38396C3-345A-4209-9675-1C0CD95E3134")]
	[ClassInterface(ClassInterfaceType.None)]
	[ComVisible(true)]
	public class Calc : MarshalByRefObject, ICalc
    {
		public IDelphiEventContainer DelphiEvents { get; set; }
		public int Sum(int a, int b)
		{
			// Pokud je naplnena property pouzije se callback do delphi a vrati se jeho vysledek
			// .. coz je (a + b) * 100
			// Jinak se vrati pouze a + b
			if (DelphiEvents != null)
			{
				int[] paramDoDelphi = new int[] { a, b };
				int delphiResult = DelphiEvents.GetValue(paramDoDelphi);
				return delphiResult;
			}

			// Defaultni navrat pokud neni naplnen delphi callback
			return a + b;
		}
	}

	// Toto rozhrani se implementuje v delphi
	// TDelphiEventsContainer = class(TInterfacedObject, IDelphiEventContainer)
	[Guid("FB3C8BAA-27FF-4E8C-89D6-133F86EC47F1")]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]    // U delphi dedene tridy InterfaceIsunknown
	[ComVisible(true)]
	public interface IDelphiEventContainer
	{
		int GetValue(int[] parValue);
	}

	[Guid("071CD8E4-4FDD-4919-8817-AD5133B7682A")]
	[ClassInterface(ClassInterfaceType.None)]  
	[ComVisible(true)]
	public class DelphiEventContainer : MarshalByRefObject, IDelphiEventContainer
	{
		// metoda, ktera se v delphi prepise ( nevrati -1 )
		public int GetValue(int[] parValue)		
		{
			return -1;
		}
	}
}
