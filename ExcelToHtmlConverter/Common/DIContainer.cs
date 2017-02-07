using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToHtmlConverter.Common
{
    public class DIContainer
    {
        private static DIContainer instance = new DIContainer();

        public static DIContainer Instance
        {
            get { return instance; }
        }

        private DIContainer()
        {

        }

        public delegate object Creator(DIContainer container);

        private readonly Dictionary<string, object> configuration = new Dictionary<string, object>();
        private readonly Dictionary<Type, Creator> typeToCreator = new Dictionary<Type, Creator>();

        public Dictionary<string, object> Configuration
        {
            get { return configuration; }
        }

        public void Register<T>(Creator creator)
        {
            typeToCreator.Add(typeof(T), creator);
        }

        public T Create<T>()
        {
            var typekey = typeof(T);
            if (!typeToCreator.ContainsKey(typekey))
            {
                TryAutoRegisterTypeFromAppDomain(typekey);
            }

            if (!typeToCreator.ContainsKey(typekey))
            {
                TryAutoRegisterFromAssembliesInWorkingDir(typekey);
            }
            return (T)typeToCreator[typekey](this);

        }

        private void TryAutoRegisterFromAssembliesInWorkingDir(Type type)
        {
            var directory = Directory.GetCurrentDirectory();
            var dlls = new DirectoryInfo(directory).EnumerateFiles("*.dll", SearchOption.TopDirectoryOnly);

            foreach (var dll in dlls)
            {
                try
                {
                    var assembly = Assembly.LoadFile(dll.FullName);
                    var concrete = assembly.GetTypes()
                                    .Where(t => type.IsAssignableFrom(t) && t.IsInterface == false)
                                    .FirstOrDefault();
                    if (concrete != null)
                    {
                        typeToCreator.Add(type, delegate { return Activator.CreateInstance(concrete); });
                        break;
                    }
                }
                catch (Exception e)
                {
                    // This assembly get skipped
                }
            }
        }

        public void TryAutoRegisterTypeFromAppDomain(Type type)
        {
            var concrete = AppDomain.CurrentDomain.GetAssemblies()
                .ToList()
                .SelectMany(a => a.GetTypes())
                .Where(t => type.IsAssignableFrom(t) && t.IsInterface == false)
                .FirstOrDefault();

            if (concrete != null)
            {
                typeToCreator.Add(type, delegate { return Activator.CreateInstance(concrete); });
            }

        }

        public T GetConfiguration<T>(string name)
        {
            return (T)configuration[name];
        }

        public void Reset()
        {
            configuration.Clear();
            typeToCreator.Clear();
        }
    }
}
