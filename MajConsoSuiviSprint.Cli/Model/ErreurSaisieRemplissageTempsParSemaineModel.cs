using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MajConsoSuiviSprint.Cli.Model
{
    internal class ErreurSaisieRemplissageTempsParSemaineModel
    {   
        public string Qui { get; set; } = default!;
        public int NumeroDeSemaine { get; set; } = default!;
        public float HeureDeclaree { get; set; } = default!;
    }
}
