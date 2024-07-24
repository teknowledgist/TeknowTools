## Summary
   Determine if the processor has specific features/instruction sets.

## DESCRIPTION
It is not often that a PowerShell developer needs to know about features of
the processor running the code.  If an application's binaries are compiled 
to take advantage of specific features of a processor though, This function 
can help identify and report back on the availability of those features.  
It does not identify the operating system's ability to use those features.

To read about the features identified, [see Microsoft's descriptions here](https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-isprocessorfeaturepresent).

## Parameters
This function has no input parameters

## OUTPUT
   Returns a hashtable of feature names and a boolean for their availability.

## Examples
```
Prompt> Get-ProcessorFeatures()

Name                           Value
----                           -----
COMPARE_EXCHANGE128            True
SSE4_2_INSTRUCTIONS            True
ARM_64BIT_LOADSTORE_ATOMIC     False
ARM_V83_JSCVT_INSTRUCTIONS     False
NX_ENABLED                     True
ARM_V8_CRC32_INSTRUCTIONS      False
SSE4_1_INSTRUCTIONS            True
CHANNELS_ENABLED               False
ARM_V8_CRYPTO_INSTRUCTIONS     False
AVX_INSTRUCTIONS               True
SECOND_LEVEL_ADDRESS_TRANSL... True
RDWRFSGSBASE                   True
ARM_V81_ATOMIC_INSTRUCTIONS    False
ARM_V82_DP_INSTRUCTIONS        False
XMMI_INSTRUCTIONS              True
FLOATING_POINT_PRECISION_ER... False
FASTFAIL                       True
SSSE3_INSTRUCTIONS             True
AVX2_INSTRUCTIONS              True
ARM_VFP_32_REGISTERS           False
ARM_V8_INSTRUCTIONS            False
RDTSC_INSTRUCTION              True
ARM_V83_LRCPC_INSTRUCTIONS     False
SSE3_INSTRUCTIONS              True
COMPARE64_EXCHANGE128          False
ARM_DIVIDE_INSTRUCTION         False
COMPARE_EXCHANGE_DOUBLE        True
ARM_FMAC_INSTRUCTIONS          False
VIRT_FIRMWARE_ENABLED          True
3DNOW_INSTRUCTIONS             False
FLOATING_POINT_EMULATED        False
AVX512F_INSTRUCTIONS           False
XMMI64_INSTRUCTIONS            True
XSAVE_ENABLED                  True
PAE_ENABLED                    True
ARM_EXTERNAL_CACHE             False
MMX_INSTRUCTIONS               True
```

```
$Features = Get-ProcessorFeatures()
if ($Features["AVX512F_INSTRUCTIONS"]) {
   Write-Host "This processor has AVX512 features."
}
```

## License
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
