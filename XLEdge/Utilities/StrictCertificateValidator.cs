using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Security;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace XLEdge.Utilities
{
    public static class StrictCertificateValidator
    {
        public static bool Validate(
            object sender,
            X509Certificate certificate,
            X509Chain chain,
            SslPolicyErrors sslPolicyErrors)
        {
            try
            {
                var cert2 = certificate as X509Certificate2
                            ?? new X509Certificate2(certificate);

                LogCertificate(cert2, chain, sslPolicyErrors);

                // 🔒 Absolute rule: ANY SSL policy error = FAIL
                if (sslPolicyErrors != SslPolicyErrors.None)
                {
                    LogUtility.LogError($"TLS validation failed: {sslPolicyErrors}");
                    return false;
                }

                // 🔒 Force Windows chain validation with revocation
                using (var strictChain = new X509Chain())
                {
                    strictChain.ChainPolicy = new X509ChainPolicy
                    {
                        RevocationMode = X509RevocationMode.Online,
                        RevocationFlag = X509RevocationFlag.ExcludeRoot,
                        VerificationFlags = X509VerificationFlags.NoFlag,
                        UrlRetrievalTimeout = TimeSpan.FromSeconds(15)
                    };

                    strictChain.ChainPolicy.ApplicationPolicy.Add(
                        new Oid("1.3.6.1.5.5.7.3.1")); // Server Authentication

                    if (!strictChain.Build(cert2))
                    {
                        foreach (var status in strictChain.ChainStatus)
                        {
                            LogUtility.LogError(
                                $"Chain validation error: {status.Status} - {status.StatusInformation}");
                        }

                        return false;
                    }
                }

                // 🔒 Enforce strong crypto (defense-in-depth)
                if (!IsStrongCertificate(cert2))
                {
                    LogUtility.LogError("Certificate rejected: weak signature or key length");
                    return false;
                }

                string url = TryGetRequestTarget(sender);
                LogUtility.LogDebug("TLS certificate validation successful for " + url);
                return true;
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Fatal TLS validation exception: {ex}");
                return false;
            }
        }
        private static string TryGetRequestTarget(object sender)
        {
            try
            {
                return sender switch
                {
                    HttpRequestMessage httpRequest => httpRequest.RequestUri?.ToString(),
                    HttpWebRequest webRequest => webRequest.RequestUri?.ToString(),
                    _ => $"Unknown (sender type: {sender?.GetType().FullName})",
                };
            }
            catch (Exception ex)
            {
                // Safe to ignore: best-effort request-target lookup, used only for a diagnostic log
                // line; falls back to an empty string.
                LogUtility.LogDebug($"{nameof(TryGetRequestTarget)}: failed to resolve request target for logging - {ex.Message}");
                return string.Empty;
            }

        }
        private static bool IsStrongCertificate(X509Certificate2 cert)
        {
            // cert.PublicKey.Key is the old, RSA/DSA-only accessor - it throws NotSupportedException
            // for ECDSA keys, which is exactly what most present-day TLS certificates use (Google,
            // Wikipedia/Cloudflare, and most modern CDNs default to ECDSA). That exception was
            // escaping all the way out to Validate()'s outer catch, logging "Fatal TLS validation
            // exception" and failing otherwise-legitimate, trusted connections. GetRSAPublicKey()/
            // GetECDsaPublicKey() are the modern, algorithm-specific replacements - they return null
            // instead of throwing when the certificate's key isn't that algorithm.
            if (cert.GetRSAPublicKey() is System.Security.Cryptography.RSA rsa)
            {
                // RSA < 2048 bits → reject
                if (rsa.KeySize < 2048)
                {
                    return false;
                }
            }
            else if (cert.GetECDsaPublicKey() is System.Security.Cryptography.ECDsa ecdsa)
            {
                // ECDSA < 256 bits (i.e. weaker than P-256) → reject. P-256 is the accepted modern
                // minimum and is roughly equivalent in strength to RSA-2048.
                if (ecdsa.KeySize < 256)
                {
                    return false;
                }
            }
            else
            {
                // Neither RSA nor ECDSA (e.g. legacy DSA) - not a key type this method can size up,
                // so treat it the same as a weak key rather than letting it through unchecked.
                LogUtility.LogError($"Certificate rejected: unrecognized public key algorithm ({cert.PublicKey.Oid.FriendlyName})");
                return false;
            }

            // Reject weak signature algorithms
            string sigAlg = cert.SignatureAlgorithm.FriendlyName;
            if (sigAlg != null &&
                (sigAlg.IndexOf("md5", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 sigAlg.IndexOf("sha1", StringComparison.OrdinalIgnoreCase) >= 0))
            {
                return false;
            }

            return true;
        }

        private static void LogCertificate(
            X509Certificate2 cert,
            X509Chain chain,
            SslPolicyErrors errors)
        {
            // Consolidated into one line (was 7 separate LogDebug calls plus banner lines) -
            // still captures every field useful for diagnosing a rejected/unexpected certificate.
            LogUtility.LogDebug($"TLS certificate validation: Subject={cert.Subject}, Issuer={cert.Issuer}, Thumbprint={cert.Thumbprint}, Valid={cert.NotBefore:u} to {cert.NotAfter:u}, SslPolicyErrors={errors}");

            if (chain?.ChainStatus != null)
            {
                foreach (var status in chain.ChainStatus)
                {
                    LogUtility.LogDebug(
                        $"Chain status: {status.Status} - {status.StatusInformation}");
                }
            }
        }
    }
}
