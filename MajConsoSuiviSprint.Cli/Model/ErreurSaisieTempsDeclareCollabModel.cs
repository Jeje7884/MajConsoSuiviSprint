using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MajConsoSuiviSprint.Cli.Model
{
    internal class ErreurSaisieTempsDeclareCollabModel
    {
        public string TrigrammeCollab { get; set; } = default!;
        public string NumeroDeSemaine { get; set; }= default!;
        public float    NombreHeurDeclaree { get; set; }
    }
}
