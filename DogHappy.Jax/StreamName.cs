using System.IO;

namespace DogHappy.Jax
{
    public class StreamName<T> where T : Stream
    {
        public T Stream { get; set; }
        public string Name { get; set; }
    }
}
