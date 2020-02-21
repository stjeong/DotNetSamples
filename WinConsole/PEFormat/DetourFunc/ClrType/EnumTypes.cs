using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DetourFunc.Clr
{
    public enum MethodClassification
    {
        mcIL = 0, // IL
        mcFCall = 1, // FCall (also includes tlbimped ctor, Delegate ctor)
        mcNDirect = 2, // N/Direct
        mcEEImpl = 3, // special method; implementation provided by EE (like Delegate Invoke)
        mcArray = 4, // Array ECall
        mcInstantiated = 5, // Instantiated generic methods, including descriptors
                            // for both shared and unshared code (see InstantiatedMethodDesc)

        mcComInterop = 6,
        mcDynamic = 7, // for method dewsc with no metadata behind
        mcCount,
    };

    public enum MethodKind
    {
        KindMask = 0x07,
        GenericMethodDefinition = 0x00,
        UnsharedMethodInstantiation = 0x01,
        SharedMethodInstantiation = 0x02,
        WrapperStubWithInstantiations = 0x03,

#if EnC_SUPPORTED
    // Non-virtual method added through EditAndContinue.
    EnCAddedMethod = 0x07,
#endif // EnC_SUPPORTED

        Unrestored = 0x08,

#if FEATURE_COMINTEROP
    HasComPlusCallInfo = 0x10, // this IMD contains an optional ComPlusCallInfo
#endif // FEATURE_COMINTEROP
    }

    public enum MethodDescFlags3
    {
        // There are flags available for use here (currently 5 flags bits are available); however, new bits are hard to come by, so any new flags bits should
        // have a fairly strong justification for existence.
        TokenRemainderMask = 0x3FFF, // This must equal METHOD_TOKEN_REMAINDER_MASK calculated higher in this file
                                     // These are seperate to allow the flags space available and used to be obvious here
                                     // and for the logic that splits the token to be algorithmically generated based on the
                                     // #define
        HasForwardedValuetypeParameter = 0x4000, // Indicates that a type-forwarded type is used as a valuetype parameter (this flag is only valid for ngenned items)

        ValueTypeParametersWalked = 0x4000, // Indicates that all typeref's in the signature of the method have been resolved to typedefs (or that process failed) (this flag is only valid for non-ngenned methods)
        DoesNotHaveEquivalentValuetypeParameters = 0x8000, // Indicates that we have verified that there are no equivalent valuetype parameters for this method
        /*
    #ifdef FEATURE_TYPEEQUIVALENCE
        inline BOOL DoesNotHaveEquivalentValuetypeParameters()
        {
            LIMITED_METHOD_DAC_CONTRACT;
            return (m_wFlags3AndTokenRemainder & enum_flag3_DoesNotHaveEquivalentValuetypeParameters) != 0;
        }

        inline void SetDoesNotHaveEquivalentValuetypeParameters()
        {
            LIMITED_METHOD_CONTRACT;
            InterlockedUpdateFlags3(enum_flag3_DoesNotHaveEquivalentValuetypeParameters, TRUE);
        }
    #endif // FEATURE_TYPEEQUIVALENCE
        */
    };

    public enum PackedSlotLayout
    {
        SlotMask = 0x03FF,
        NameHashMask = 0xFC00
    };

    public enum CorTokenType
    {
        mdtModule = 0x00000000,       //
        mdtTypeRef = 0x01000000,       //
        mdtTypeDef = 0x02000000,       //
        mdtFieldDef = 0x04000000,       //
        mdtMethodDef = 0x06000000,       //
        mdtParamDef = 0x08000000,       //
        mdtInterfaceImpl = 0x09000000,       //
        mdtMemberRef = 0x0a000000,       //
        mdtCustomAttribute = 0x0c000000,       //
        mdtPermission = 0x0e000000,       //
        mdtSignature = 0x11000000,       //
        mdtEvent = 0x14000000,       //
        mdtProperty = 0x17000000,       //
        mdtMethodImpl = 0x19000000,       //
        mdtModuleRef = 0x1a000000,       //
        mdtTypeSpec = 0x1b000000,       //
        mdtAssembly = 0x20000000,       //
        mdtAssemblyRef = 0x23000000,       //
        mdtFile = 0x26000000,       //
        mdtExportedType = 0x27000000,       //
        mdtManifestResource = 0x28000000,       //
        mdtGenericParam = 0x2a000000,       //
        mdtMethodSpec = 0x2b000000,       //
        mdtGenericParamConstraint = 0x2c000000,

        mdtString = 0x70000000,       //
        mdtName = 0x71000000,       //
        mdtBaseType = 0x72000000,       // Leave this on the high end value. This does not correspond to metadata table
    }

    [Flags]
    public enum MethodDescFlags2
    {
        // enum_flag2_HasPrecode implies that enum_flag2_HasStableEntryPoint is set.
        HasStableEntryPoint = 0x01,   // The method entrypoint is stable (either precode or actual code)
        HasPrecode = 0x02,   // Precode has been allocated for this method

        IsUnboxingStub = 0x04,
        HasNativeCodeSlot = 0x08,   // Has slot for native code

        IsJitIntrinsic = 0x10,   // Jit may expand method as an intrinsic

        IsEligibleForTieredCompilation = 0x20,

        // unused                           = 0x40,
        // unused                           = 0x80,
    }

    public enum MethodDescClassification
    {
        // Method is IL, FCall etc., see MethodClassification above.
        mdcClassification = 0x0007,
        mdcClassificationCount = mdcClassification + 1,

        // Note that layout of code:MethodDesc::s_ClassificationSizeTable depends on the exact values
        // of mdcHasNonVtableSlot and mdcMethodImpl

        // Has local slot (vs. has real slot in MethodTable)
        mdcHasNonVtableSlot = 0x0008,

        // Method is a body for a method impl (MI_MethodDesc, MI_NDirectMethodDesc, etc)
        // where the function explicitly implements IInterface.foo() instead of foo().
        mdcMethodImpl = 0x0010,

        // Method is static
        mdcStatic = 0x0020,

        // unused                           = 0x0040,
        // unused                           = 0x0080,
        // unused                           = 0x0100,
        // unused                           = 0x0200,

        // Duplicate method. When a method needs to be placed in multiple slots in the
        // method table, because it could not be packed into one slot. For eg, a method
        // providing implementation for two interfaces, MethodImpl, etc
        mdcDuplicate = 0x0400,

        // Has this method been verified?
        mdcVerifiedState = 0x0800,

        // Is the method verifiable? It needs to be verified first to determine this
        mdcVerifiable = 0x1000,

        // Is this method ineligible for inlining?
        mdcNotInline = 0x2000,

        // Is the method synchronized
        mdcSynchronized = 0x4000,

        // Does the method's slot number require all 16 bits
        mdcRequiresFullSlotNumber = 0x8000
    };

    [Flags]
    public enum WFLAGS_HIGH_ENUM : uint
    {
        // DO NOT use flags that have bits set in the low 2 bytes.
        // These flags are DWORD sized so that our atomic masking
        // operations can operate on the entire 4-byte aligned DWORD
        // instead of the logical non-aligned WORD of flags.  The
        // low WORD of flags is reserved for the component size.

        // The following bits describe mutually exclusive locations of the type
        // in the type hiearchy.
        enum_flag_Category_Mask = 0x000F0000,

        enum_flag_Category_Class = 0x00000000,
        enum_flag_Category_Unused_1 = 0x00010000,
        enum_flag_Category_Unused_2 = 0x00020000,
        enum_flag_Category_Unused_3 = 0x00030000,

        enum_flag_Category_ValueType = 0x00040000,
        enum_flag_Category_ValueType_Mask = 0x000C0000,
        enum_flag_Category_Nullable = 0x00050000, // sub-category of ValueType
        enum_flag_Category_PrimitiveValueType = 0x00060000, // sub-category of ValueType, Enum or primitive value type
        enum_flag_Category_TruePrimitive = 0x00070000, // sub-category of ValueType, Primitive (ELEMENT_TYPE_I, etc.)

        enum_flag_Category_Array = 0x00080000,
        enum_flag_Category_Array_Mask = 0x000C0000,
        // enum_flag_Category_IfArrayThenUnused                 = 0x00010000, // sub-category of Array
        enum_flag_Category_IfArrayThenSzArray = 0x00020000, // sub-category of Array

        enum_flag_Category_Interface = 0x000C0000,
        enum_flag_Category_Unused_4 = 0x000D0000,
        enum_flag_Category_Unused_5 = 0x000E0000,
        enum_flag_Category_Unused_6 = 0x000F0000,

        enum_flag_Category_ElementTypeMask = 0x000E0000, // bits that matter for element type mask


        enum_flag_HasFinalizer = 0x00100000, // instances require finalization

        enum_flag_IfNotInterfaceThenMarshalable = 0x00200000, // Is this type marshalable by the pinvoke marshalling layer
#if FEATURE_COMINTEROP
    enum_flag_IfInterfaceThenHasGuidInfo = 0x00200000, // Does the type has optional GuidInfo
#endif // FEATURE_COMINTEROP

        enum_flag_ICastable = 0x00400000, // class implements ICastable interface

        enum_flag_HasIndirectParent = 0x00800000, // m_pParentMethodTable has double indirection

        enum_flag_ContainsPointers = 0x01000000,

        enum_flag_HasTypeEquivalence = 0x02000000, // can be equivalent to another type

#if FEATURE_COMINTEROP
    enum_flag_HasRCWPerTypeData = 0x04000000, // has optional pointer to RCWPerTypeData
#endif // FEATURE_COMINTEROP

        enum_flag_HasCriticalFinalizer = 0x08000000, // finalizer must be run on Appdomain Unload
        enum_flag_Collectible = 0x10000000,
        enum_flag_ContainsGenericVariables = 0x20000000,   // we cache this flag to help detect these efficiently and
                                                           // to detect this condition when restoring

        enum_flag_ComObject = 0x40000000, // class is a com object

        enum_flag_HasComponentSize = 0x80000000,   // This is set if component size is used for flags.

        // Types that require non-trivial interface cast have this bit set in the category
        enum_flag_NonTrivialInterfaceCast = enum_flag_Category_Array
                                             | enum_flag_ComObject
                                             | enum_flag_ICastable

    };  // enum WFLAGS_HIGH_ENUM

    [Flags]
    public enum WFLAGS_LOW_ENUM
    {
        // AS YOU ADD NEW FLAGS PLEASE CONSIDER WHETHER Generics::NewInstantiation NEEDS
        // TO BE UPDATED IN ORDER TO ENSURE THAT METHODTABLES DUPLICATED FOR GENERIC INSTANTIATIONS
        // CARRY THE CORECT FLAGS.
        //

        // We are overloading the low 2 bytes of m_dwFlags to be a component size for Strings
        // and Arrays and some set of flags which we can be assured are of a specified state
        // for Strings / Arrays, currently these will be a bunch of generics flags which don't
        // apply to Strings / Arrays.

        enum_flag_UNUSED_ComponentSize_1 = 0x00000001,

        enum_flag_StaticsMask = 0x00000006,
        enum_flag_StaticsMask_NonDynamic = 0x00000000,
        enum_flag_StaticsMask_Dynamic = 0x00000002,   // dynamic statics (EnC, reflection.emit)
        enum_flag_StaticsMask_Generics = 0x00000004,   // generics statics
        enum_flag_StaticsMask_CrossModuleGenerics = 0x00000006, // cross module generics statics (NGen)
        enum_flag_StaticsMask_IfGenericsThenCrossModule = 0x00000002, // helper constant to get rid of unnecessary check

        enum_flag_NotInPZM = 0x00000008,   // True if this type is not in its PreferredZapModule

        enum_flag_GenericsMask = 0x00000030,
        enum_flag_GenericsMask_NonGeneric = 0x00000000,   // no instantiation
        enum_flag_GenericsMask_GenericInst = 0x00000010,   // regular instantiation, e.g. List<String>
        enum_flag_GenericsMask_SharedInst = 0x00000020,   // shared instantiation, e.g. List<__Canon> or List<MyValueType<__Canon>>
        enum_flag_GenericsMask_TypicalInst = 0x00000030,   // the type instantiated at its formal parameters, e.g. List<T>

        enum_flag_HasVariance = 0x00000100,   // This is an instantiated type some of whose type parameters are co- or contra-variant

        enum_flag_HasDefaultCtor = 0x00000200,
        enum_flag_HasPreciseInitCctors = 0x00000400,   // Do we need to run class constructors at allocation time? (Not perf important, could be moved to EEClass

#if FEATURE_HFA
#if UNIX_AMD64_ABI
#error "Can't define both FEATURE_HFA and UNIX_AMD64_ABI"
#endif
        enum_flag_IsHFA                     = 0x00000800,   // This type is an HFA (Homogenous Floating-point Aggregate)
#endif // FEATURE_HFA

#if UNIX_AMD64_ABI
#if FEATURE_HFA
#error "Can't define both FEATURE_HFA and UNIX_AMD64_ABI"
#endif
        enum_flag_IsRegStructPassed         = 0x00000800,   // This type is a System V register passed struct.
#endif // UNIX_AMD64_ABI

        enum_flag_IsByRefLike = 0x00001000,

        // In a perfect world we would fill these flags using other flags that we already have
        // which have a constant value for something which has a component size.
        enum_flag_UNUSED_ComponentSize_5 = 0x00002000,
        enum_flag_UNUSED_ComponentSize_6 = 0x00004000,
        enum_flag_UNUSED_ComponentSize_7 = 0x00008000,

        //#define SET_FALSE(flag)     ((flag) & 0)
        //#define SET_TRUE(flag)      ((flag) & 0xffff)

        // IMPORTANT! IMPORTANT! IMPORTANT!
        //
        // As you change the flags in WFLAGS_LOW_ENUM you also need to change this
        // to be up to date to reflect the default values of those flags for the
        // case where this MethodTable is for a String or Array
        enum_flag_StringArrayValues = (enum_flag_StaticsMask_NonDynamic & 0xffff) |
                                      (enum_flag_NotInPZM & 0) |
                                      (enum_flag_GenericsMask_NonGeneric & 0xffff) |
                                      (enum_flag_HasVariance & 0) |
                                      (enum_flag_HasDefaultCtor & 0) |
                                      (enum_flag_HasPreciseInitCctors & 0),

    };  // enum WFLAGS_LOW_ENUM

    [Flags]
    public enum WFLAGS2_ENUM
    {
        // AS YOU ADD NEW FLAGS PLEASE CONSIDER WHETHER Generics::NewInstantiation NEEDS
        // TO BE UPDATED IN ORDER TO ENSURE THAT METHODTABLES DUPLICATED FOR GENERIC INSTANTIATIONS
        // CARRY THE CORECT FLAGS.

        // The following bits describe usage of optional slots. They have to stay
        // together because of we index using them into offset arrays.
        enum_flag_MultipurposeSlotsMask = 0x001F,
        enum_flag_HasPerInstInfo = 0x0001,
        enum_flag_HasInterfaceMap = 0x0002,
        enum_flag_HasDispatchMapSlot = 0x0004,
        enum_flag_HasNonVirtualSlots = 0x0008,
        enum_flag_HasModuleOverride = 0x0010,

        enum_flag_IsZapped = 0x0020, // This could be fetched from m_pLoaderModule if we run out of flags

        enum_flag_IsPreRestored = 0x0040, // Class does not need restore
                                          // This flag is set only for NGENed classes (IsZapped is true)

        enum_flag_HasModuleDependencies = 0x0080,

        enum_flag_IsIntrinsicType = 0x0100,

        enum_flag_RequiresDispatchTokenFat = 0x0200,

        enum_flag_HasCctor = 0x0400,
        enum_flag_HasCCWTemplate = 0x0800, // Has an extra field pointing to a CCW template

#if FEATURE_64BIT_ALIGNMENT
    enum_flag_RequiresAlign8 = 0x1000, // Type requires 8-byte alignment (only set on platforms that require this and don't get it implicitly)
#endif

        enum_flag_HasBoxedRegularStatics = 0x2000, // GetNumBoxedRegularStatics() != 0

        enum_flag_HasSingleNonVirtualSlot = 0x4000,

        enum_flag_DependsOnEquivalentOrForwardedStructs = 0x8000, // Declares methods that have type equivalent or type forwarded structures in their signature

    };  // enum WFLAGS2_ENUM
}
