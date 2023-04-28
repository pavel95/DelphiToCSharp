using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace SharpToDelphiCallback
{
	// Declaration of delegates- these are callbacks to delphi, use of MarshalAs is recommended
	public delegate void DelphiIntArrayCallback([MarshalAs(UnmanagedType.SafeArray)] int[] ids);
	public delegate void DelphiSomethinElseCallback([MarshalAs(UnmanagedType.Interface)] ISomethingElse somethingElse);

	// Interface
	// This is needed
	// .. GUID
	// .. Atribut [InterfaceType(ComInterfaceType.InterfaceIsDual)]
	// .. Atribut [ComVisible(true)]
	[Guid("79801B97-62BC-489B-935A-8543116E620E")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	[ComVisible(true)]
	public interface ISomething
	{
		void SetIntArrayCallback(int integerPointer);
		void SetSomethingElseCallback(int integerPointer);

		void Reverse([MarshalAs(UnmanagedType.SafeArray)] int[] ids);
		void CallSomethingElse([MarshalAs(UnmanagedType.Interface)] ISomethingElse somethingElse);
	}

	// .. Implementation of interface, class must inherite from MarshalRefObject
	[Guid("B2BB5506-8308-400C-924F-4A3E3D46D965")]
	[ClassInterface(ClassInterfaceType.None)]
	[ComVisible(true)]
	public class Something : MarshalByRefObject, ISomething
	{
		DelphiIntArrayCallback delphiIntArrayCallback = null;
		DelphiSomethinElseCallback delphiSomethingElseCallback = null;

		// This reverse array that come from delphi and send it back through callback
		// Array comes forth and back as argument
		public void Reverse([MarshalAs(UnmanagedType.SafeArray)] int[] ids)
		{
			if (delphiIntArrayCallback != null)
			{
				try
				{
					ids.Reverse();
					delphiIntArrayCallback(ids);
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.Message);
				}
			}
		}

		// This will use an object of type somethingElse and then send it back through callback
		// object comes forth and back as argument
		public void CallSomethingElse([MarshalAs(UnmanagedType.Interface)] ISomethingElse somethingElse)
		{
			//SomethingElse somethingElse = new SomethingElse();
			somethingElse.ShowMessage("Hi, I am in the C# dll now!");
			delphiSomethingElseCallback(somethingElse);
		}
		

		// Set the array method callback to delphi
		public void SetIntArrayCallback(int integerPointer)
		{
			try
			{
				IntPtr intPtr = new IntPtr(integerPointer);
				delphiIntArrayCallback = (DelphiIntArrayCallback)Marshal.GetDelegateForFunctionPointer(intPtr, typeof(DelphiIntArrayCallback));
			}
			catch (Exception e)
			{
				MessageBox.Show(e.Message);
			}
		}

		// set the object method callback to delphi
		public void SetSomethingElseCallback(int integerPointer)
		{
			try
			{
				IntPtr intPtr = new IntPtr(integerPointer);
				delphiSomethingElseCallback = (DelphiSomethinElseCallback)Marshal.GetDelegateForFunctionPointer(intPtr, typeof(DelphiSomethinElseCallback));
			}
			catch (Exception e)
			{
				MessageBox.Show(e.Message);
			}
		}
	}

	// Interface and its implementation for testing purpouses - sending it from delphi to C# and back to delphi 

	[Guid("00A7D870-AE65-4682-BD1B-D2EE21B0B88C")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	[ComVisible(true)]
	public interface ISomethingElse
	{
		void ShowMessage(string message);
	}

	[Guid("165B7823-3B02-4DAA-8205-C07FB7FA124A")]
	[ClassInterface(ClassInterfaceType.None)]
	[ComVisible(true)]
	public class SomethingElse : MarshalByRefObject, ISomethingElse
	{
		public SomethingElse()
		{

		}
		public void ShowMessage(string message)
		{
			MessageBox.Show(message);
		}
	}
}
