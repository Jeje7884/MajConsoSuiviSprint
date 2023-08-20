start-transcript -path  "c:\temp\cd13\test.log" | out-null;
if ($args.length -ge 1)
  {
    Write-Host "args"
   
      $filenameconfig= $args[0];
   
  }
 
  Write-Host "nbe param $($args.length)"
 

#   [PSCustomObject]$users = $jsonobject
 
 write-host "---------zz----";
 write-host $filenameconfig 
 write-Output $filenameconfig | Out-File .\Utils\fileJson.json



# ;
  # $test="{2022:{Application:SRS,NumeroDeDemande:TS-2022,HeureTotaleDeDeveloppement:10.0,HeureTotaleDeQualification:10.0,IsDemandeValide:false},2025:{Application:SRS,NumeroDeDemande:TS-2025,HeureTotaleDeDeveloppement:10.0,HeureTotaleDeQualification:10.0,IsDemandeValide:false}}"
  
   $JsonContent = Get-Content .\Utils\fileJson.json | Out-String | ConvertFrom-Json

  # Write-Host "Processing JSON object:"
  # Write-Host "Property1: $($jsonObject.Property1)"
  # Write-Host "Property2: $($jsonObject.Property2)"