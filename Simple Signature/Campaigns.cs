﻿// Generated by Xamasoft JSON Class Generator
// http://www.xamasoft.com/json-class-generator

using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Simple_Signature
{

    public class Signatures
    {

        [JsonProperty("_id")]
        public string Id { get; set; }

        [JsonProperty("createdAt")]
        public DateTime CreatedAt { get; set; }

        [JsonProperty("firm")]
        public string Firm { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("value")]
        public string Value { get; set; }

        [JsonProperty("img")]
        public string Image { get; set; }

        static public Signatures getDefaultInterne(Signatures[] campaigns, SimpleSign s)
        {
            foreach (var item in campaigns)
            {
                if(item.Name == "interne default")
                {
                    s.currentSignature = item;
                    return item;
                }
            }
            return null;
        }

        static public Signatures getDefaultExterne(Signatures[] campaigns, SimpleSign s)
        {
            foreach (var item in campaigns)
            {
                if (item.Name == "externe default")
                {
                    s.currentSignature = item;
                    return item;
                }
            }
            return getDefaultInterne(campaigns, s);
        }

        static public Signatures getFirstExterne(Signatures[] campaigns, SimpleSign s)
        {
            foreach (var item in campaigns)
            {
                if (item.Name != "externe default" && item.Name != "interne default")
                {
                    s.currentSignature = item;
                    return item;
                }
            }
            return getDefaultExterne(campaigns, s);
        }
    }

}

