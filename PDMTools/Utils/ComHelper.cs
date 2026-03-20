using System;
using System.Runtime.InteropServices;

namespace PDMTools.Utils
{
    internal static class ComHelper
    {
        public static void Release(object comObject)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                if (Marshal.IsComObject(comObject))
                {
                    Marshal.FinalReleaseComObject(comObject);
                }
            }
            catch
            {
                // 釋放失敗不應阻斷主流程
            }
        }
    }
}
