using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;

namespace ExcelExtruder
{
    public class TypeConvert
    {
        public Action<string> _EVENT_ERROR_LOG;
        private Assembly m_assembly;
        private Dictionary<Type, MethodInfo> m_tryParseMethodInfos;
        private Dictionary<string, Type> m_foundType;

        public static Dictionary<Type, Delegate> TryParseDelegates = new Dictionary<Type, Delegate>();

        /// <summary>
        /// 内置类型解析器（零反射），覆盖所有常见值类型
        /// 返回解析后的 object，解析失败返回 null
        /// </summary>
        private static readonly Dictionary<Type, Func<string, object>> s_builtinParsers
            = new Dictionary<Type, Func<string, object>>
        {
            { typeof(int),     s => int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : null },
            { typeof(uint),    s => uint.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : null },
            { typeof(long),    s => long.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : null },
            { typeof(ulong),   s => ulong.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : null },
            { typeof(short),   s => short.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : null },
            { typeof(ushort),  s => ushort.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : null },
            { typeof(byte),    s => byte.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : null },
            { typeof(sbyte),   s => sbyte.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : null },
            { typeof(float),   s => float.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var v) ? v : null },
            { typeof(double),  s => double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var v) ? v : null },
            { typeof(decimal), s => decimal.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out var v) ? v : null },
            { typeof(bool),    s => bool.TryParse(s, out var v) ? v : null },
            { typeof(char),    s => s.Length == 1 ? s[0] : null },
        };

        public string m_currentPlacement;
        public uint m_currentPlacementId;
        public int m_currentArrayIndex = 0;

        public string m_extraInfo = "";

        public TypeConvert(Action<string> EVENT_ERROR_LOG)
        {
            _EVENT_ERROR_LOG = EVENT_ERROR_LOG;
        }

        public void Init(Assembly assembly)
        {
            m_assembly = assembly;
            m_tryParseMethodInfos = new Dictionary<Type, MethodInfo>();
            m_foundType = new Dictionary<string, Type>(StringComparer.Ordinal);
        }

        private void Error(string str)
        {
            _EVENT_ERROR_LOG?.Invoke(str + "; " + m_extraInfo);
        }

        public Type TryGetType(string typeName)
        {
            if (string.IsNullOrEmpty(typeName))
                return null;

            if (m_foundType.TryGetValue(typeName, out var cached))
                return cached;

            Type t = null;

            // 尝试从动态程序集中获取类型
            if (m_assembly != null)
                t = FindTypeInCustomAssembly(typeName);

            // 尝试从默认程序集中获取类型
            if (t == null)
                t = Assembly.GetExecutingAssembly().GetType(typeName);

            // 尝试从当前已加载程序集中获取完整类型名
            if (t == null)
            {
                foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
                {
                    t = assembly.GetType(typeName, false);
                    if (t != null)
                        break;
                }
            }

            // 尝试在当前已加载程序集中按简单类型名匹配
            if (t == null)
                t = FindUniqueTypeBySimpleName(AppDomain.CurrentDomain.GetAssemblies(), typeName);

            // 尝试从系统库中获取类型
            if (t == null)
                t = Type.GetType(typeName);

            // 如果是枚举，需要获取+后面的类型名称
            if (t == null)
                t = TryGetEnumType(typeName);

            if (t == null)
            {
                Error("[TryGetType] Can't find the type: " + typeName);
                return null;
            }
            else
            {
                m_foundType[typeName] = t;
                return t;
            }
        }

        private Type FindTypeInCustomAssembly(string typename)
        {
            if (m_assembly == null)
                return null;

            var type = m_assembly.GetType(typename, false);
            if (type != null)
                return type;

            return FindUniqueTypeBySimpleName(new[] { m_assembly }, typename);
        }

        private Type TryGetEnumType(string typeName)
        {
            // aaaa+bbbb
            var index = typeName.IndexOf('+');
            if (index != -1)
            {
                var subName = typeName.Substring(index + 1);
                return TryGetType(subName);
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// 动态创建指定类型的数组（无反射）
        /// </summary>
        public static Array CreateArray(Type innerType, int length)
        {
            return Array.CreateInstance(innerType, length);
        }

        /// <summary>
        /// 动态创建Generic
        /// 创建Dictionary请使用CreateDictionary
        /// </summary>
        /// <param name="generic"></param>
        /// <param name="innerType"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        public static object CreateGeneric(Type generic, Type innerType, params object[] args)
        {
            System.Type specificType = generic.MakeGenericType(new Type[] { innerType });
            return Activator.CreateInstance(specificType, args);
        }

        /// <summary>
        /// 动态创建Dictionary
        /// </summary>
        /// <param name="keyType"></param>
        /// <param name="valueType"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        public static object CreateDictionary(Type keyType, Type valueType, params object[] args)
        {
            Type genericType = typeof(Dictionary<,>);
            Type specificType = genericType.MakeGenericType(new Type[] { keyType, valueType });
            return Activator.CreateInstance(specificType, args);
        }

        public bool TryParse(Type type, string value, out object result)
        {
            return TryParse(type, value, -1, out result);
        }

        /// <summary>
        /// 将字符串转换为真实的类型
        /// </summary>
        /// <param name="type"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public bool TryParse(Type type, string value, int arraySize, out object result)
        {
            if (string.IsNullOrEmpty(value))
            {
                // value为空，直接返回空值
                result = null;
                return true;
            }
            else if (type == typeof(string))
            {
                // 是字符串，直接返回
                result = Trim(value);
                return true;
            }
            else if (type.IsEnum)
            {
                // 枚举的情况
                return TryParse2Enum(type, value, out result);
            }
            else if (type.IsArray)
            {
                // 数组的情况
                var ret = TryParse2Array(type, value, out result);
                if (ret && arraySize != -1)
                {
                    Array ar = result as Array;
                    if (ar.Length != arraySize)
                    {
                        Error($"The size of the array dose not meet the definition length {arraySize}: {value}");
                        return false;
                    }
                }
                return ret;
            }
            else if (type.IsGenericType)
            {
                // 是GenericType的情况
                return TryParse2Generic(type, value, out result);
            }
            else if (TryParse2Custom(type, value, out var flag, out result))
            {
                return flag;
            }
            else
            {
                try
                {
                    return TryParse2Object(type, value, out result);
                }
                catch (Exception ex)
                {
                    Error("[TryParse] " + ex.Message);
                    UnityEngine.Debug.LogException(ex);
                    result = null;
                    return false;
                }
            }
        }

        protected virtual bool TryParse2Custom(Type type, string value, out bool flag, out object result)
        {
            result = default;
            flag = false;
            return false;
        }

        private bool TryParse2Object(Type type, string value, out object result)
        {
            var trimmed = Trim(value);

            // ① 内置类型快速路径（零反射）
            if (s_builtinParsers.TryGetValue(type, out var parser))
            {
                result = parser(trimmed);
                if (result != null) return true;
                Error($"[TryParse2Builtin] 解析失败: [{value}] => {type.FullName}");
                return false;
            }

            // ② 自定义类型：通过缓存的 MethodInfo 调用 TryParse
            MethodInfo mi = GetMethodInfo(type, "TryParse");
            if (mi != null)
            {
                var parameters = new object[] { trimmed, GetDefaultValue(type) };
                if ((bool)mi.Invoke(null, parameters) == true)
                {
                    result = parameters[1];
                    return true;
                }
                else
                {
                    Error(string.Format("[TryParse2Object] TryParse method return fail: [{1}] => {0}", type.FullName, value));
                    result = null;
                    return false;
                }
            }

            // ③ 自定义类型：通过注册的委托调用 TryParse
            if (TryParseDelegates.TryGetValue(type, out var del))
            {
                var parameters = new object[] { trimmed, GetDefaultValue(type) };
                if ((bool)del.DynamicInvoke(parameters) == true)
                {
                    result = parameters[1];
                    return true;
                }
                else
                {
                    Error(string.Format("[TryParse2Object] TryParse delegate return fail: [{1}] => {0}", type.FullName, value));
                    result = null;
                    return false;
                }
            }

            // ④ 无法解析
            Error("[TryParse2Object] Can't find the TryParse method of type: " + type.FullName);
            result = null;
            return false;
        }

        /// <summary>
        /// 获取某个type的方法
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        private MethodInfo GetMethodInfo(Type type, string methodName)
        {
            MethodInfo result;
            if (!m_tryParseMethodInfos.TryGetValue(type, out result))
            {
                MethodInfo mi = type.GetMethod(methodName, BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy, Type.DefaultBinder,
                                    new Type[] { typeof(string), type.MakeByRefType() },
                                    new ParameterModifier[] { new ParameterModifier(2) });
                if (mi != null)
                    m_tryParseMethodInfos.Add(type, mi);

                return mi;
            }
            else
            {
                return result;
            }
        }

        private bool TryParse2Generic(Type type, string value, out object result)
        {
            var cutName = GetGenericTypeRealName(type);

            switch (cutName)
            {
                case "List": return Parse2Generic(type, typeof(List<>), "Add", value, out result);
                case "HashSet": return Parse2Generic(type, typeof(HashSet<>), "Add", value, out result);
                case "Queue": return Parse2Generic(type, typeof(Queue<>), "Enqueue", value, out result);
                case "Stack": return Parse2Generic(type, typeof(Stack<>), "Push", value, out result);
                case "LinkedList": return Parse2LinkedList(type, value, out result);
                case "Dictionary": return Parse2Dictionary(type, value, out result);
                default:
                    // 暂不支持其他Generic类型
                    Error("[TryParse2Generic] Don't support the generic type: " + type.FullName);
                    result = null;
                    return false;
            }
        }

        private bool Parse2Generic(Type type, Type genericType, string addMethodName, string value, out object result)
        {
            var innerType = type.GenericTypeArguments[0];
            result = CreateGeneric(genericType, innerType);

            // 对实现 IList 的集合（如 List<T>），直接使用接口调用，无需反射
            if (result is IList list)
                return PushInner(value, innerType, item => list.Add(item));

            // 其他集合（HashSet/Queue/Stack）仍需反射获取方法
            var addMethodInfo = result.GetType().GetMethod(addMethodName);
            var collection = result;
            return PushInner(value, innerType, item => addMethodInfo.Invoke(collection, new[] { item }));
        }

        private bool Parse2LinkedList(Type type, string value, out object result)
        {
            var innerType = type.GenericTypeArguments[0];
            result = CreateGeneric(typeof(LinkedList<>), innerType);

            var allMethod = result.GetType().GetMethods();
            foreach (var method in allMethod)
            {
                if (method.Name == "AddLast" && method.ReturnType.FullName.IndexOf("LinkedListNode") != -1)
                {
                    var linkedList = result;
                    var addLastMethod = method;
                    return PushInner(value, innerType, item => addLastMethod.Invoke(linkedList, new[] { item }));
                }
            }

            result = null;
            return false;
        }

        /// <summary>
        /// 向集合中逐个添加解析后的元素
        /// 使用 Action 委托替代 MethodInfo.Invoke，调用方决定添加方式
        /// </summary>
        private bool PushInner(string str, Type innerType, Action<object> addAction)
        {
            bool allSuccess = true;

            var tempstring = Trim(str);
            string[] strs = CutStringByGroup(tempstring);
            for (int i = 0; i < strs.Length; i++)
            {
                m_currentArrayIndex = i;
                object innerResult = null;
                if (TryParse(innerType, strs[i], out innerResult))
                    addAction(innerResult);
                else
                    allSuccess = false;
            }

            return allSuccess;
        }

        private bool Parse2Dictionary(Type type, string value, out object result)
        {
            var keyType = type.GenericTypeArguments[0];
            var valueType = type.GenericTypeArguments[1];

            object generic = CreateDictionary(keyType, valueType);
            // 使用 IDictionary 接口直接添加，无需反射
            var dict = (IDictionary)generic;

            var groups = CutStringByGroup(Trim(value));
            bool allSuccess = true;
            for (int i = 0; i < groups.Length; i++)
            {
                var str = groups[i];

                int index = str.IndexOf(":");
                if (index == -1)
                {
                    Error("[TryParse2Dictionary] Can't convert value to dictionary group: " + str);
                    allSuccess = false;
                    continue;
                }

                string keyStr = str.Substring(0, index);
                string valueStr = str.Substring(index + 1);

                if ((valueType.IsGenericType || valueType.IsArray) && valueStr.StartsWith("[") && valueStr.EndsWith("]"))
                    valueStr = valueStr.Substring(1, valueStr.Length - 2);

                object keyResult = null;
                object valueResult = null;

                if (TryParse(keyType, keyStr, out keyResult) && TryParse(valueType, valueStr, out valueResult))
                    dict.Add(keyResult, valueResult);
                else
                    allSuccess = false;
            }

            result = generic;
            return allSuccess;
        }

        private bool TryParse2Array(Type type, string value, out object result)
        {
            var innerType = type.GetElementType();

            string[] strs = CutStringByGroup(value);
            // 直接使用 Array.CreateInstance + Array.SetValue，无需反射
            Array array = CreateArray(innerType, strs.Length);

            bool allSuccess = true;
            for (int i = 0; i < strs.Length; i++)
            {
                m_currentArrayIndex = i;
                object innerResult = null;
                if (TryParse(innerType, strs[i], out innerResult))
                    array.SetValue(innerResult, i);
                else
                    allSuccess = false;
            }

            result = array;
            return allSuccess;
        }

        protected virtual char SplitChar => ';';
        protected virtual string[] CutStringByGroup(string str)
        {
            List<string> result = new List<string>();
            if (string.IsNullOrWhiteSpace(str))
                return result.ToArray();

            int bracketsDeep = 0;
            int startIndex = 0;

            for (int i = 0; i < str.Length; i++)
            {
                var cha = str[i];
                if (cha == '[')
                    bracketsDeep += 1;

                if (cha == ']')
                    bracketsDeep -= 1;

                if (cha == SplitChar && bracketsDeep == 0)
                {
                    result.Add(NormalizeGroupToken(str.Substring(startIndex, i - startIndex)));
                    startIndex = i + 1;
                }
            }

            result.Add(NormalizeGroupToken(str.Substring(startIndex)));
            return result.ToArray();
        }

        private bool TryParse2Enum(Type type, string value, out object result)
        {
            try
            {
                result = Enum.Parse(type, value);
                return true;
            }
            catch (Exception ex)
            {
                UnityEngine.Debug.LogError($"解析失败。请检查Excel表{m_extraInfo}");
                result = default;
                Error("[TryParse2Enum] " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// 获取真实字符串
        /// Excel保存为Uncode文本时有可能在字符串首位加入引号，需要手动去除
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string Trim(string value)
        {
            string str = value.Trim();
            if (str.Length > 2
                && str.Split('\"').Length == 3
                && str.StartsWith("\"")
                && str.EndsWith("\""))
                str = str.Substring(1, str.Length - 2);

            return str;
        }

        /// <summary>
        /// 从泛型类型中提取不带 arity 后缀的类型名
        /// 例如: List`1 → List, Dictionary`2 → Dictionary
        /// </summary>
        public static string GetGenericTypeRealName(Type type)
        {
            var name = type.Name; // 例如: "List`1", "Dictionary`2"
            int index = name.IndexOf('`');
            return index >= 0 ? name.Substring(0, index) : name;
        }

        private static object GetDefaultValue(Type type)
        {
            return type.IsValueType ? Activator.CreateInstance(type) : null;
        }

        private static string NormalizeGroupToken(string token)
        {
            var buff = token.Trim();
            if (buff.StartsWith("[") && buff.EndsWith("]"))
                buff = buff.Substring(1, buff.Length - 2).Trim();

            return buff;
        }

        private Type FindUniqueTypeBySimpleName(IEnumerable<Assembly> assemblies, string typeName)
        {
            List<Type> matches = null;

            foreach (var assembly in assemblies)
            {
                if (assembly == null)
                    continue;

                foreach (var type in GetLoadableTypes(assembly))
                {
                    if (!string.Equals(type.Name, typeName, StringComparison.Ordinal))
                        continue;

                    matches ??= new List<Type>();
                    if (!matches.Contains(type))
                        matches.Add(type);
                }
            }

            if (matches == null || matches.Count == 0)
                return null;

            if (matches.Count > 1)
            {
                Error("[TryGetType] Ambiguous type name: " + typeName + ". Matches: " + string.Join(", ", matches.ConvertAll(t => t.FullName)));
                return null;
            }

            return matches[0];
        }

        private static IEnumerable<Type> GetLoadableTypes(Assembly assembly)
        {
            try
            {
                return assembly.GetTypes();
            }
            catch (ReflectionTypeLoadException ex)
            {
                var result = new List<Type>();
                foreach (var type in ex.Types)
                {
                    if (type != null)
                        result.Add(type);
                }
                return result;
            }
        }
    }
}
