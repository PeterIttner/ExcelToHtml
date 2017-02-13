using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelToHtmlConverter.Common
{
    /// <summary>
    /// Dependency Injection Container
    /// </summary>
    public class DIContainer
    {
        /// <summary>
        /// Delegate that is used for registering methods that create instances of objects
        /// </summary>
        /// <param name="container">The Dependency Injection container that is used</param>
        /// <returns></returns>
        public delegate object Creator(DIContainer container);

        private readonly Dictionary<string, object> configuration = new Dictionary<string, object>();
        private readonly Dictionary<Type, Creator> typeToCreator = new Dictionary<Type, Creator>();
        private static DIContainer instance = new DIContainer();

        #region C'tor

        /// <summary>
        /// Gets the single instance of the DIContainer
        /// </summary>
        public static DIContainer Instance
        {
            get { return instance; }
        }

        private DIContainer()
        {

        }

        #endregion

        #region Public Interface

        /// <summary>
        /// Registers a new creates method for creating instances of objects
        /// </summary>
        /// <typeparam name="T">The type that should be created</typeparam>
        /// <param name="creator">The method that creates the actual instance</param>
        public void Register<T>(Creator creator)
        {
            typeToCreator.Add(typeof(T), creator);
        }

        /// <summary>
        /// Try to create an instance of the given type.
        /// Throws an exception if the instance could not be created.
        /// If the type is not registered, the current assembly and all 
        /// assemblies in the current directory will be scanned for the given type.
        /// </summary>
        /// <typeparam name="T">The type that should be created</typeparam>
        /// <returns>An instance of the given type, if the creation is possible</returns>
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

        /// <summary>
        /// Gets the Configuration of the container.
        /// Can be used to register objects with strings.
        /// </summary>
        public Dictionary<string, object> Configuration
        {
            get { return configuration; }
        }

        /// <summary>
        /// Gets an instance that is registered with a configured name.
        /// Throws an exception if the instance could not be found.
        /// </summary>
        /// <typeparam name="T">The type of the instance that sould be returned</typeparam>
        /// <param name="name">The name with which the instance has be registered</param>
        /// <returns>The registered instance if any.</returns>
        public T GetConfiguration<T>(string name)
        {
            return (T)configuration[name];
        }

        /// <summary>
        /// Resets all registered data.
        /// </summary>
        public void Reset()
        {
            configuration.Clear();
            typeToCreator.Clear();
        }

        #endregion

        #region Private Helper Methods

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

        private void TryAutoRegisterTypeFromAppDomain(Type type)
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

        #endregion
    }
}
