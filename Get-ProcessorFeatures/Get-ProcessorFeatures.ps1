<#
.SYNOPSIS
   Determine if the processor has specific features/instruction sets.
.DESCRIPTION
   Applications binaries may be compiled to take advantage of specific features
   of a processor.  This function can identify and report back on the 
   availability of those features.  It does not identify whether the 
   operating system can take advantage of those features.

   To read about the features identified, see here:
   https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-isprocessorfeaturepresent
.OUTPUT
   Returns a hashtable of feature names and a boolean for their availability.
.Example
   $Features = Get-ProcessorFeatures()
   if ($Features["AVX512F_INSTRUCTIONS"]) {
      Write-Host "This processor has AVX512 features."
   }
.NOTES
   Copyright 2024 Teknowledgist

   This script/information is free: you can redistribute 
   it and/or modify it under the terms of the GNU General Public License 
   as published by the Free Software Foundation, either version 2 of the 
   License, or (at your option) any later version.

   This script is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   The GNU General Public License can be found at <http://www.gnu.org/licenses/>.
#>

# https://www.p-invoke.net/kernel32/isprocessorfeaturepresent
$Signature = @'
[DllImport("Kernel32.dll")][return: MarshalAs(UnmanagedType.Bool)]
public static extern bool IsProcessorFeaturePresent(
   uint ProcessorFeature
);
'@

$type = Add-Type -MemberDefinition $Signature -Name Win32Utils -Namespace GetProcessorFeatures -PassThru
   
# https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-isprocessorfeaturepresent
$FeatureList = @{
   '0' = 'FLOATING_POINT_PRECISION_ERRATA' #On a Pentium, a floating-point precision error can occur in rare circumstances.   
   '1' = 'FLOATING_POINT_EMULATED' #Floating-point operations are emulated using a software emulator.   
   '2' = 'COMPARE_EXCHANGE_DOUBLE' #The atomic compare and exchange operation (cmpxchg) is available.   
   '3' = 'MMX_INSTRUCTIONS' #The MMX instruction set is available.   
   '6' = 'XMMI_INSTRUCTIONS' #The SSE instruction set is available.   
   '7' = '3DNOW_INSTRUCTIONS' #The 3D-Now instruction set is available.   
   '8' = 'RDTSC_INSTRUCTION' #The RDTSC instruction is available.   
   '9' = 'PAE_ENABLED' #The processor is Physical Address Extension (PAE)-enabled. All x64 processors always return a nonzero value for this feature.    
   '10' = 'XMMI64_INSTRUCTIONS' #The SSE2 instruction set is available. Windows 2000: Not supported.    
   '12' = 'NX_ENABLED' #Data execution prevention is enabled.  Windows XP/2000: Not supported.    
   '13' = 'SSE3_INSTRUCTIONS' #The SSE3 instruction set is available. Windows Server 2003 and Windows XP/2000: Not supported.    
   '14' = 'COMPARE_EXCHANGE128' #The  atomic compare and exchange 128-bit operation (cmpxchg16b) is available. Windows Server 2003 and Windows XP/2000: Not supported.    
   '15' = 'COMPARE64_EXCHANGE128' #The atomic compare 64 and exchange 128-bit operation (cmp8xchg16) is available. Windows Server 2003 and Windows XP/2000: Not supported.    
   '16' = 'CHANNELS_ENABLED' #The processor channels are enabled.   
   '17' = 'XSAVE_ENABLED' #The processor implements the XSAVE and XRSTOR instructions. Windows Server 2003/2008, Windows 2000/XP/Vista: Not supported.    
   '18' = 'ARM_VFP_32_REGISTERS' #The VFP/Neon: 32 x 64bit register bank is present. This flag has the same meaning as PF_ARM_VFP_EXTENDED_REGISTERS .   
   '20' = 'SECOND_LEVEL_ADDRESS_TRANSLATION' #Second Level Address Translation is supported by the hardware.   
   '21' = 'VIRT_FIRMWARE_ENABLED' #Virtualization is enabled in the firmware and made available by the operating system.   
   '22' = 'RDWRFSGSBASE' #RDFSBASE, RDGSBASE, WRFSBASE, and WRGSBASE instructions are available.   
   '23' = 'FASTFAIL' #_fastfail() is available.   
   '24' = 'ARM_DIVIDE_INSTRUCTION' #The divide instructions are available.   
   '25' = 'ARM_64BIT_LOADSTORE_ATOMIC' #The 64-bit load/store atomic instructions are available.   
   '26' = 'ARM_EXTERNAL_CACHE' #The external cache is available.   
   '27' = 'ARM_FMAC_INSTRUCTIONS' #The floating-point multiply-accumulate instruction is available.   
   '29' = 'ARM_V8_INSTRUCTIONS' #This Arm processor implements the Arm v8 instructions set.   
   '30' = 'ARM_V8_CRYPTO_INSTRUCTIONS' #This Arm processor implements the Arm v8 extra cryptographic instructions (for example, AES, SHA1 and SHA2).   
   '31' = 'ARM_V8_CRC32_INSTRUCTIONS' #This Arm processor implements the Arm v8 extra CRC32 instructions.   
   '34' = 'ARM_V81_ATOMIC_INSTRUCTIONS' #This Arm processor implements the Arm v8.1 atomic instructions (for example, CAS, SWP).   
   '36' = 'SSSE3_INSTRUCTIONS' #The SSSE3 instruction set is available.   
   '37' = 'SSE4_1_INSTRUCTIONS' #The SSE4_1 instruction set is available.   
   '38' = 'SSE4_2_INSTRUCTIONS' #The SSE4_2 instruction set is available.   
   '39' = 'AVX_INSTRUCTIONS' #The AVX instruction set is available.   
   '40' = 'AVX2_INSTRUCTIONS' #The AVX2 instruction set is available.   
   '41' = 'AVX512F_INSTRUCTIONS' #The AVX512F instruction set is available.   
   '43' = 'ARM_V82_DP_INSTRUCTIONS' #This Arm processor implements the Arm v8.2 DP instructions (for example, SDOT, UDOT). This feature is optional in Arm v8.2 implementations and mandatory in Arm v8.4 implementations.   
   '44' = 'ARM_V83_JSCVT_INSTRUCTIONS' #This Arm processor implements the Arm v8.3 JSCVT instructions (for example, FJCVTZS).   
   '45' = 'ARM_V83_LRCPC_INSTRUCTIONS' #This Arm processor implements the Arm v8.3 LRCPC instructions (for example, LDAPR). Note that certain Arm v8.2 CPUs may optionally support the LRCPC instructions.   
}

Foreach ($Feature in $FeatureList.keys) {
   @{$FeatureList["$Feature"] = $type::IsProcessorFeaturePresent($Feature)}
}




