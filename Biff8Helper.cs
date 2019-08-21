using System;
using System.Collections.Generic;
using System.IO ;
using System.Text;

namespace AFC.WorkStation.ExcelReport
{
	public class Biff8Helper
	{
		// BIFF Operation 
		/// <summary>
		/// Seek for BOF 809H.
		/// </summary>
		/// <param name="fp"></param>
		/// <returns></returns>
		public static int GotoBOF(FileStream fp)
		{
			byte[] buf = new byte[2];
			int len = 0;
			while ((len = fp.Read(buf, 0, 2)) == 2)
			{
				// 809h BOF.
				if (buf[0] == 0x09 && buf[1] == 0x08)
				{
					fp.Seek(-2, SeekOrigin.Current);

					return (int)fp.Position;
				}
			}
			return -1; // EOF 

		}
		
		public static BiffRecord NextBiff (FileStream fp)
		{
			byte[] buf = new byte[4];
			int len = fp.Read (buf, 0, 4) ;
			
			if (len != 4)
				return null ;
				
			BiffRecord r = new BiffRecord ();

			r.opCode = BitConverter.ToUInt16 (buf, 0) ;
			r.len = BitConverter.ToUInt16(buf, 2);
			
			r.info = new byte[r.len];
			
			len = fp.Read (r.info, 0, r.len) ;
			if (len != r.len)
				return null ;
			
			return r ;
		}
		
		public static int RemoveWarning (FileStream fp)
		{
			if (GotoBOF (fp) < 0)
				return -1 ;
			BiffRecord r = null ;
			while ((r = NextBiff (fp)) != null)
			{
				long p = fp.Position ;
				switch (r.opCode)
				{
				case 0x3D : // WINDOW1 
					// 4 * 2 Wnd pos 
					// grbit 
					fp.Seek (-(r.len - 4*2 - 2), SeekOrigin.Current) ;
					// Update itabCur to 0.
					fp.Write (new byte [2], 0, 2);
					fp.Flush ();
					fp.Seek (p, SeekOrigin.Begin) ;
					break ;
				case 0x85 : // BOUNDSHEET 
					
					string sheetName = ExtractUnicodeString (r.info, 6) ;
					if (sheetName.IndexOf("Warning") >= 0)
					{
						// mark this Sheet as Very Hidden
						fp.Seek (- (r.len + 4), SeekOrigin.Current) ;
						
						fp.Write (new byte[] { 0x02} , 0, 1);

						fp.Seek(p, SeekOrigin.Begin);
					}
					break ;	
				case 0X0A : // EOF
					return 0 ;
				}
			}
			return -1 ;
		}

		private static string ExtractUnicodeString (byte[] info, int offset)
		{
			int cch = info[offset];
			byte cchGrbit = info[offset + 1];
			
			if (cchGrbit == 0) // fHighByte compressed string 
			{
				// as default ;
				return Encoding.Default.GetString (info, offset + 2, cch) ;
			}
			else 
				// as UNICODE 
			{
				return Encoding.Unicode.GetString (info, offset + 2, cch*2) ;
			}
				
		}
	}
	
	public class BiffRecord
	{
		public ushort opCode ;
		public ushort len ;
		public byte [] info ;
	}
	
}
