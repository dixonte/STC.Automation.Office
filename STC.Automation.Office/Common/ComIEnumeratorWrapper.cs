using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.Runtime.InteropServices;
using System.Reflection;

namespace STC.Automation.Office.Common
{
    internal class ComIEnumeratorWrapper<T> : IEnumerator<T>
        where T: ComWrapper
    {
        private IEnumerator _iEnum;

        internal ComIEnumeratorWrapper(object enumVariant)
        {
            if (!(enumVariant is IEnumerator))
            {
                throw new COMException("Problem wrapping IEnumVARIANT object; does not support interface IEnumerator.");
            }

            _iEnum = enumVariant as IEnumerator;
        }

        #region IEnumerator Members

        public object Current
        {
            get { return this.Current; }
        }

        public bool MoveNext()
        {
            return _iEnum.MoveNext();
        }

        public void Reset()
        {
            _iEnum.Reset();
        }

        #endregion

        #region IEnumerator<T> Members

        T IEnumerator<T>.Current
        {
            get
            {
                T currentWrapper = (T)typeof(T).GetConstructor(BindingFlags.NonPublic | BindingFlags.Instance, null, new Type[] { typeof(object) }, null).Invoke(new object[] { _iEnum.Current });

                return currentWrapper;
            }
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            // Nothing to do?
        }

        #endregion
    }
}
