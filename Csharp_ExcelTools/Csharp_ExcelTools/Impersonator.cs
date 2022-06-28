public class Impersonator : System.IDisposable
{
    //登入提供者
    protected const int LOGON32_PROVIDER_DEFAULT = 0;
    protected const int LOGON32_LOGON_INTERACTIVE = 2;
    public WindowsIdentity Identity = null;
    private System.IntPtr m_accessToken;
    [System.Runtime.InteropServices.DllImport("advapi32.dll", SetLastError = true)]
    private static extern bool LogonUser(string lpszUsername, string lpszDomain,
    string lpszPassword, int dwLogonType, int dwLogonProvider, ref System.IntPtr phToken);
    [System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
    private extern static bool CloseHandle(System.IntPtr handle);
    public Impersonator()
    {
        //建構子
    }
    public void Login(string username, string domain, string password)
    {
        if (this.Identity != null)
        {
            this.Identity.Dispose();
            this.Identity = null;
        }
        try
        {
            this.m_accessToken = new System.IntPtr(0);
            Logout();
            this.m_accessToken = System.IntPtr.Zero;
            //執行LogonUser
            bool isSuccess = LogonUser(
               username,
               domain,
               password,
               LOGON32_LOGON_INTERACTIVE,
               LOGON32_PROVIDER_DEFAULT,
               ref this.m_accessToken);

            if (!isSuccess)
            {
                //取得錯誤碼
                int error = System.Runtime.InteropServices.Marshal.GetLastWin32Error();
                throw new System.ComponentModel.Win32Exception(error);
            }
            Identity = new WindowsIdentity(this.m_accessToken);
        }
        catch
        {
            throw;
        }
    }
    //使用完之後登出，放在Dispose中呼叫
    public void Logout()
    {
        if (this.m_accessToken != System.IntPtr.Zero)
            CloseHandle(m_accessToken);
        this.m_accessToken = System.IntPtr.Zero;
        if (this.Identity != null)
        {
            this.Identity.Dispose();
            this.Identity = null;
        }
    }
    void System.IDisposable.Dispose()
    {
        Logout();
    }
}


