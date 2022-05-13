using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Animation;
using MathNet.Numerics;
using OxyPlot;

namespace WPF_DCSOperations
{



    public class Info : Observable
    {
        public Info() { }

        public string FileName { get; set; } = "";


        private string _CompanyName;
        public string CompanyName
        {
            get { return _CompanyName; }
            set
            {
                if (value != _CompanyName)
                {
                    _CompanyName = value;
                    NotifyChange("");
                }
            }
        }


        private DateTime _StartDate;
        public DateTime StartDate
        {
            get { return _StartDate; }
            set
            {
                if (value != _StartDate)
                {
                    _StartDate = value;
                    NotifyChange("");
                }
            }
        }


        private DateTime _EndDate;
        public DateTime EndDate
        {
            get { return _EndDate; }
            set
            {
                if (value != _EndDate)
                {
                    _EndDate = value;
                    NotifyChange("");
                }
            }
        }


        private string _IMONumber;
        public string IMONumber
        {
            get { return _IMONumber; }
            set
            {
                if (value != _IMONumber)
                {
                    _IMONumber = value;
                    NotifyChange("");
                }
            }
        }


        private string _ShipType;
        public string ShipType
        {
            get { return _ShipType; }
            set
            {
                if (value != _ShipType)
                {
                    _ShipType = value;
                    NotifyChange("");
                }
            }
        }


        private string _VesselSize;
        public string VesselSize
        {
            get
            {
                if (ShipType == "BULK CARRIER")
                {
                    if (DeadWeight >= 0 && DeadWeight <= 9999) return _VesselSize = "1: 0-9999";
                    if (DeadWeight >= 10000 && DeadWeight <= 34999) return _VesselSize = "2: 10000-34999";
                    if (DeadWeight >= 35000 && DeadWeight <= 59999) return _VesselSize = "3: 35000-59999";
                    if (DeadWeight >= 60000 && DeadWeight <= 99999) return _VesselSize = "4: 60000-99999";
                    if (DeadWeight >= 100000 && DeadWeight <= 199999) return _VesselSize = "5: 100000-199999";
                    if (DeadWeight >= 200000) return _VesselSize = "6: 200000-+"; ;
                }

                if (ShipType == "CHEMICAL TANKER")
                {
                    if (DeadWeight >= 0 && DeadWeight <= 4999) return _VesselSize = "1: 0-4999";
                    if (DeadWeight >= 5000 && DeadWeight <= 9999) return _VesselSize = "2: 5000-9999";
                    if (DeadWeight >= 10000 && DeadWeight <= 19999) return _VesselSize = "3: 10000-19999";
                    if (DeadWeight >= 20000) return _VesselSize = "4: 20000-+"; ;
                }

                if (ShipType == "CONTAINER")
                {
                    if (DeadWeight >= 0 && DeadWeight <= 999) return _VesselSize = "1: 0-999";
                    if (DeadWeight >= 1000 && DeadWeight <= 1999) return _VesselSize = "2: 1000-1999";
                    if (DeadWeight >= 2000 && DeadWeight <= 2999) return _VesselSize = "3: 2000-2999";
                    if (DeadWeight >= 3000 && DeadWeight <= 4999) return _VesselSize = "4: 3000-4999";
                    if (DeadWeight >= 5000 && DeadWeight <= 7999) return _VesselSize = "5: 5000-7999";
                    if (DeadWeight >= 8000 && DeadWeight <= 11999) return _VesselSize = "6: 8000-11999";
                    if (DeadWeight >= 12000 && DeadWeight <= 14500) return _VesselSize = "7: 12000-14500";
                    if (DeadWeight >= 20000) return _VesselSize = "8: 20000-+"; ;
                }

                if (ShipType == "GENERAL CARGO")
                {
                    if (DeadWeight >= 0 && DeadWeight <= 4999) return _VesselSize = "1: 0-4999";
                    if (DeadWeight >= 5000 && DeadWeight <= 9999) return _VesselSize = "2: 5000-9999";
                    if (DeadWeight >= 10000) return _VesselSize = "3: 10000-+"; ;
                }

                if (ShipType == "LIQUEFIED GAS TANKER")
                {
                    if (DeadWeight >= 0 && DeadWeight <= 49999) return _VesselSize = "1: 0-49999";
                    if (DeadWeight >= 50000 && DeadWeight <= 199999) return _VesselSize = "2: 50000-199999";
                    if (DeadWeight >= 200000) return _VesselSize = "3: 200000-+"; ;
                }

                if (ShipType == "OIL TANKER")
                {
                    if (DeadWeight >= 0 && DeadWeight <= 4999) return _VesselSize = "1: 0-4999";
                    if (DeadWeight >= 5000 && DeadWeight <= 9999) return _VesselSize = "2: 5000-9999";
                    if (DeadWeight >= 10000 && DeadWeight <= 19999) return _VesselSize = "3: 10000-19999";
                    if (DeadWeight >= 20000 && DeadWeight <= 59999) return _VesselSize = "4: 20000-59999";
                    if (DeadWeight >= 60000 && DeadWeight <= 79999) return _VesselSize = "5: 60000-79999";
                    if (DeadWeight >= 80000 && DeadWeight <= 119999) return _VesselSize = "6: 80000-119999";
                    if (DeadWeight >= 120000 && DeadWeight <= 199999) return _VesselSize = "7: 120000-199999";
                    if (DeadWeight >= 200000) return _VesselSize = "8: 200000-+"; ;
                }

                else
                {
                    return _VesselSize = "Other";
                }
                return _VesselSize = "Other";
            }
            set
            {
                if (value != _VesselSize)
                {
                    _VesselSize = value;
                    NotifyChange("");
                }
            }

        }


        private int _GrossTonnage;
        public int GrossTonnage
        {
            get { return _GrossTonnage; }
            set
            {
                if (value != _GrossTonnage)
                {
                    _GrossTonnage = value;
                    NotifyChange("");
                }
            }
        }


        private int _NetTonnage;
        public int NetTonnage
        {
            get { return _NetTonnage; }
            set
            {
                if (value != _NetTonnage)
                {
                    _NetTonnage = value;
                    NotifyChange("");
                }
            }
        }


        private double _DeadWeight;
        public double DeadWeight
        {
            get { return _DeadWeight; }
            set
            {
                if (value != _DeadWeight)
                {
                    _DeadWeight = value;
                    NotifyChange("");
                }
            }
        }


        private double _EEDI;
        public double EEDI
        {
            get { return _EEDI; }
            set
            {
                if (value != _EEDI)
                {
                    _EEDI = value;
                    NotifyChange("");
                }
            }
        }


        private string _ICEClass;
        public string ICEClass
        {
            get { return _ICEClass; }
            set
            {
                if (value != _ICEClass)
                {
                    _ICEClass = value;
                    NotifyChange("");
                }
            }
        }


        private double _MPPower;
        public double MPPower
        {
            get { return _MPPower; }
            set
            {
                if (value != _MPPower)
                {
                    _MPPower = value;
                    NotifyChange("");
                }
            }
        }


        private double _EPPower;
        public double EPPower
        {
            get { return _EPPower; }
            set
            {
                if (value != _EPPower)
                {
                    _EPPower = value;
                    NotifyChange("");
                }
            }
        }


        private double _DistanceTraveled;
        public double DistanceTraveled
        {
            get { return _DistanceTraveled; }
            set
            {
                if (value != _DistanceTraveled)
                {
                    _DistanceTraveled = value;
                    NotifyChange("");
                }
            }
        }


        private double _HoursUnderway;
        public double HoursUnderway
        {
            get { return _HoursUnderway; }
            set
            {
                if (value != _HoursUnderway)
                {
                    _HoursUnderway = value;
                    NotifyChange("");
                }
            }
        }


        private double _DO;
        public double DO
        {
            get { return _DO; }
            set
            {
                if (value != _DO)
                {
                    _DO = value;
                    NotifyChange("");
                }
            }
        }



        private double _LFO;
        public double LFO
        {
            get { return _LFO; }
            set
            {
                if (value != _LFO)
                {
                    _LFO = value;
                    NotifyChange("");
                }
            }
        }


        private double _HFO;
        public double HFO
        {
            get { return _HFO; }
            set
            {
                if (value != _HFO)
                {
                    _HFO = value;
                    NotifyChange("");
                }
            }
        }


        public double AER
        {
            get
            {
                double aer = -1;
                if (DeadWeight * DistanceTraveled >= 0)
                    aer = Math.Pow(10, 6) * ((DO * 3.206) + (LFO * 3.151) + (HFO * 3.114)) / (DeadWeight * DistanceTraveled);
                return aer;
            }
        }


        public double AERTrajectory
        {
            get
            {
                if (this.ShipType == "BULK CARRIER" && this.VesselSize == "1: 0-9999") return 26.3;
                if (this.ShipType == "BULK CARRIER" && this.VesselSize == "2: 10000-34999") return 7.0;
                if (this.ShipType == "BULK CARRIER" && this.VesselSize == "3: 35000-59999") return 4.9;
                if (this.ShipType == "BULK CARRIER" && this.VesselSize == "4: 60000-99999") return 3.9;
                if (this.ShipType == "BULK CARRIER" && this.VesselSize == "5: 100000-199999") return 2.5;
                if (this.ShipType == "BULK CARRIER" && this.VesselSize == "6: 200000-+") return 2.4;
                return 0;
            }
        }

        public double DAER
        {
            get
            {
                return AER - AERTrajectory;
            }
        }


        public double Di
        {
            get
            {
                return ((AER - AERTrajectory) / AERTrajectory) * 100;
            }

        }



    }


    public class Analysis : Observable
    {
        public double SumAER { get; set; }

    }




}
