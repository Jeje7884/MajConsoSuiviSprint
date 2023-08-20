Start-Transcript -Path  "C:\temp\cd13\test.log" | Out-Null;
if ($args.length -ge 1)
  {
    Write-Host "plus d'un paramètre"
    if ($args.length -eq 1)
    {
      $fileNameConfig= $args[0];
    }
    else {
      throw "Le nombre de paramètre passé à l'application n'est pas gérée";
    }
  }

   Write-Host "ici $($fileNameConfig["2022"].NumeroDeDemande)";
  
