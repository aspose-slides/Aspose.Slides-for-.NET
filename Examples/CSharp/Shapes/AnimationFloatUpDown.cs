using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides.Animation;

/*
This example just demontrates how aliases EffectType.FloatDown and EffectType.FloatUp will work.
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    class AnimationFloatUpDown
    {
        public static void Run()
        {
            EffectType type = EffectType.Descend;
            Console.WriteLine(type == EffectType.Descend); // Should return 'true'
            Console.WriteLine(type == EffectType.FloatDown); // Should return 'true'

            type = EffectType.FloatDown;
            Console.WriteLine(type == EffectType.Descend); // Should return 'true'
            Console.WriteLine(type == EffectType.FloatDown); // Should return 'true'

            type = EffectType.Ascend;
            Console.WriteLine(type == EffectType.Ascend); // Should return 'true'
            Console.WriteLine(type == EffectType.FloatUp); // Should return 'true'

            type = EffectType.FloatUp;
            Console.WriteLine(type == EffectType.Ascend); // Should return 'true'
            Console.WriteLine(type == EffectType.FloatUp); // Should return 'true'
        }
    }
}
