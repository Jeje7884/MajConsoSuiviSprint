using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MajConsoSuiviSprint.Cli.Model
{
    internal class TempsConsommesDemandesModel
    {
        public string NumeroDemande { get; set; } = default!;
        public float  HeureDeDeveloppement { get; private set; } = default!;
        public float HeureDeQualification { get; private set; } = default!;
        public bool IsDemandeValide { get; private set; } = default!;

    }
}
