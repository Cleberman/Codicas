import java.net.InetAddress;

public class GetMyIPAddress
{
    public static void main(String args[]) throws Exception
    {
        /*
        * "InetAddress.getLocalHost()" throws UnknownHostException: Returns
        * the address of the local host. This is achieved by retrieving the
        * name of the host from the system, then resolving that name into an
        * "InetAddress".
        *
        * Note: the resolved address may be cached for a short period of time.
        */

        InetAddress myIP = InetAddress.getLocalHost();
        
        // "getHostAdress()" returns the IP address string in textual presentation

        System.out.println("Meu endereço IP é: " + myIP.getHostAdress());     
    }
}
