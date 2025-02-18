using System;
using System.Collections.Generic;
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
            m_foundType = new Dictionary<string, Type>();
        }

        private void Error(string str)
        {
            _EVENT_ERROR_LOG?.Invoke(str + "; " + m_extraInfo);
        }

        public Type TryGetType(string typeName)
        {
            if (string.IsNullOrEmpty(typeName))
                return null;

            Type t = null;

            // 尝试从动态程序集中获取类型
            if (m_assembly != null)
                t = FindTypeInCustomAssembly(typeName);

            // 尝试从默认程序集中获取类型
            if (t == null)
                t = Assembly.GetExecutingAssembly().GetType(typeName);

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
                return t;
            }
        }

        private Type FindTypeInCustomAssembly(string typename)
        {
            Type t;

            if (!m_foundType.TryGetValue(typename, out t))
            {
                if (Environment.OSVersion.Platform == PlatformID.MacOSX)
                {
                    foreach (Type type in m_assembly.GetTypes())
                    {
                        if (string.Compare(type.Name, typename, false) != 0)
                            continue;

                        t = type;
                        m_foundType.Add(typename, t);
                        break;
                    }
                }
                else
                {
                    t = m_assembly.GetType(typename);
                    if (t != null)
                        m_foundType.Add(typename, t);
                }
            }

            return t;
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
        /// 动态创建Array
        /// </summary>
        /// <param name="innerType"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        public static object CreateArray(Type innerType, params object[] args)
        {
            Type specificType = innerType.MakeArrayType();
            return Activator.CreateInstance(specificType, args);
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
                    return true;
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
            MethodInfo mi = GetMethodInfo(type, "TryParse");
            if (mi != null)
            {
                // 有TryParse方法，直接调用
                var parameters = new object[] { Trim(value), Activator.CreateInstance(type) };
                if ((bool)mi.Invoke(null, parameters) == true)
                {
                    // 返回成功
                    result = parameters[1];
                    return true;
                }
                else
                {
                    // 返回失败
                    Error(string.Format("[TryParse2Object] TryParse method return fail: [{1}] => {0}", type.FullName, value));
                    result = null;
                    return false;
                }
            }
            else if (TryParseDelegates.TryGetValue(type, out var del))
            {
                var parameters = new object[] { Trim(value), Activator.CreateInstance(type) };
                if ((bool)del.DynamicInvoke(parameters) == true)
                {
                    // 返回成功
                    result = parameters[1];
                    return true;
                }
                else
                {
                    // 返回失败
                    Error(string.Format("[TryParse2Object] TryParse method return fail: [{1}] => {0}", type.FullName, value));
                    result = null;
                    return false;
                }
            }
            else
            {
                // 没有TryParse方法
                Error("[TryParse2Object] Can't find the TryParse method of type: " + type.FullName);
                result = null;
                return false;
            }
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
            var addMethodInfo = result.GetType().GetMethod(addMethodName);
            return PushInner(value, innerType, result, addMethodInfo);
        }

        private bool Parse2LinkedList(Type type, string value, out object result)
        {
            var innerType = type.GenericTypeArguments[0];
            result = CreateGeneric(typeof(LinkedList<>), innerType);

            var allMethod = result.GetType().GetMethods();
            foreach (var method in allMethod)
            {
                if (method.Name == "AddLast" && method.ReturnType.FullName.IndexOf("LinkedListNode") != -1)
                    return PushInner(value, innerType, result, method);
            }

            result = null;
            return false;
        }

        private bool PushInner(string str, Type innerType, object generic, MethodInfo method)
        {
            bool allSuccess = true;

            var tempstring = Trim(str);
            string[] strs = CutStringByGroup(tempstring);
            for (int i = 0; i < strs.Length; i++)
            {
                m_currentArrayIndex = i;
                object innerResult = null;
                if (TryParse(innerType, strs[i], out innerResult))
                    method.Invoke(generic, new object[] { innerResult });
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
            var addMethodInfo = generic.GetType().GetMethod("Add");

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
                    addMethodInfo.Invoke(generic, new object[] { keyResult, valueResult });
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
            object array = CreateArray(innerType, new object[] { strs.Length });
            MethodInfo setValueMethodInfo = array.GetType().GetMethod("SetValue", new Type[] { typeof(object), typeof(int) });

            bool allSuccess = true;
            for (int i = 0; i < strs.Length; i++)
            {
                m_currentArrayIndex = i;
                object innerResult = null;
                if (TryParse(innerType, strs[i], out innerResult))
                    setValueMethodInfo.Invoke(array, new object[] { innerResult, i });
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

            int bracketsDeep = 0;
            int startIndex = 0;
            int length = 0;

            for (int i = 0; i < str.Length; i++)
            {
                var cha = str[i];
                if (cha == ' ') continue;

                if (cha == '[')
                    bracketsDeep += 1;

                if (cha == ']')
                    bracketsDeep -= 1;

                if ((cha == SplitChar || i == str.Length - 1) && bracketsDeep == 0)
                {
                    if (cha != SplitChar)
                        length += 1;

                    var buff = str.Substring(startIndex, length);
                    if (buff.StartsWith("[") && buff.EndsWith("]"))
                        buff = buff.Substring(1, buff.Length - 2);

                    result.Add(buff);
                    startIndex = i + 1;
                    length = 0;
                }
                else
                {
                    length += 1;
                }
            }

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
                return true;
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

        public static string GetGenericTypeRealName(Type type)
        {
            var fullName = type.Name;
            const string head = "System.Collections.Generic.";
            var cutName = fullName.Substring(head.Length);

            int index = cutName.IndexOf('`');
            cutName = cutName.Substring(0, index);

            return cutName;
        }
    }
}
