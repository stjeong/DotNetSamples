using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace DetourFunc.ClrType
{
    public sealed class VariantMarshaller : IDisposable
    {
        [DllImport("oleaut32.dll", PreserveSig = true)]
        private static extern int VariantClear(IntPtr pObject);

        unsafe static int _varSize = sizeof(ComVariant);

        IntPtr _pAlloc;
        unsafe ComVariant* _pVariant;

        public void Dispose()
        {
            if (_pAlloc != IntPtr.Zero)
            {
                VariantClear(_pAlloc);
                Marshal.FreeHGlobal(_pAlloc);
            }
        }

        public unsafe VariantMarshaller(object obj)
        {
            _pAlloc = Marshal.AllocHGlobal(_varSize);
            Marshal.GetNativeVariantForObject(obj, _pAlloc);

            _pVariant = (ComVariant*)_pAlloc.ToPointer();
        }

        public unsafe IntPtr Variant
        {
            get { return new IntPtr(_pVariant); }
        }

        public unsafe IntPtr Unknown
        {
            get
            {
                if (_pVariant == null)
                {
                    return IntPtr.Zero;
                }

                if (_pVariant->vt != (ushort)VarEnum.VT_UNKNOWN && _pVariant->vt != (ushort)VarEnum.VT_DISPATCH)
                {
                    return IntPtr.Zero;
                }

                return _pVariant->data01;
            }
        }

        public unsafe IntPtr Dispatch
        {
            get
            {
                if (_pVariant == null)
                {
                    return IntPtr.Zero;
                }

                if (_pVariant->vt != (ushort)VarEnum.VT_DISPATCH)
                {
                    return IntPtr.Zero;
                }

                return _pVariant->data01;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        struct ComVariant
        {
            public ushort vt;
            public ushort wReserved1;
            public ushort wReserved2;
            public ushort wReserved3;
            public IntPtr data01;
            public IntPtr data02;
        }
    }
}
