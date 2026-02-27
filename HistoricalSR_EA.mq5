//+------------------------------------------------------------------+
//|                    HistoricalSR_EA.mq5                           |
//|          Historical S/R Zone Intelligence EA for MT5             |
//|   Markets: Forex, Gold, Crypto, Indices | 24/5 | 1% Risk        |
//|   Zones: H4 + D1 | Bounce + Breakout Auto-Detection             |
//|   Built with: Touch Count + Volume Cluster + Swing H/L           |
//+------------------------------------------------------------------+
#property copyright   "HistoricalSR EA"
#property version     "2.00"
#property strict

#include <Trade\Trade.mqh>
#include <Trade\PositionInfo.mqh>
#include <Trade\OrderInfo.mqh>

CTrade         Trade;
CPositionInfo  PositionInfo;

//+------------------------------------------------------------------+
//| INPUT PARAMETERS                                                  |
//+------------------------------------------------------------------+

// --- Zone Detection ---
input group "=== ZONE DETECTION ==="
input int    ZoneLookbackBars     = 0;        // 0 = full history available
input int    H4_Bars              = 2000;     // H4 bars to analyze
input int    D1_Bars              = 1000;     // D1 bars to analyze
input double ZoneWidthPips        = 20.0;     // Zone width in pips (tolerance)
input int    MinTouchCount        = 2;        // Min touches to qualify as zone
input int    ZoneStrengthMin      = 30;       // Min zone score (0-100) to trade
input int    SwingLookback        = 5;        // Bars each side for swing detection

// --- Velocity Engine ---
input group "=== VELOCITY ENGINE ==="
input int    VelocityBars         = 5;        // Bars to measure velocity
input double VelocityMinPips      = 3.0;      // Min pips/bar speed to confirm
input int    ConsecutiveBarsMin   = 2;        // Min consecutive same-direction bars

// --- Trade Brain ---
input group "=== TRADE LOGIC ==="
input double BreakoutThresholdPct = 50.0;    // % of zone width to confirm breakout
input int    ConfirmationBars     = 1;        // Confirmation candles before entry

// --- Risk Management ---
input group "=== RISK MANAGEMENT ==="
input double RiskPercent          = 1.0;      // Risk % per trade
input int    ATR_Period           = 14;       // ATR period for SL/TP
input double ATR_SL_Multiplier    = 1.5;      // SL = ATR * this multiplier
input double TP1_RR               = 1.5;      // TP1 Risk:Reward ratio
input double TP2_RR               = 3.0;      // TP2 Risk:Reward ratio
input double PartialClosePercent  = 50.0;     // % to close at TP1
input double BreakevenPips        = 20.0;     // Pips profit to move SL to BE
input double TrailingStartPips    = 30.0;     // Pips profit to start trailing
input double TrailingStepPips     = 10.0;     // Trailing step in pips
input int    MagicNumber          = 202401;   // EA Magic Number
input int    MaxSlippage          = 10;       // Max slippage points

// --- Notifications ---
input group "=== NOTIFICATIONS ==="
input bool   ShowDashboard        = true;     // Show on-chart dashboard
input bool   AlertOnZoneHit       = true;     // Alert when price hits zone
input bool   SendEmailOnTrade     = true;     // Email on trade open
input bool   SendPushOnTrade      = true;     // Push notification on trade open
input bool   LogToFile            = true;     // Log trades to file
input string LogFileName          = "HistoricalSR_Log.csv"; // Log file name

//+------------------------------------------------------------------+
//| GLOBAL STRUCTURES                                                 |
//+------------------------------------------------------------------+

struct SRZone
{
   double   price;           // Zone center price
   double   upperBound;      // Zone upper boundary
   double   lowerBound;      // Zone lower boundary
   int      touchCount;      // Number of times price touched this zone
   double   volumeScore;     // Normalized volume cluster score
   bool     isSwing;         // Is this a swing high/low
   int      strength;        // Composite strength score 0-100
   bool     isResistance;    // true = resistance, false = support
   datetime lastTouched;     // Last time price touched this zone
   bool     active;          // Zone still active (not broken)
   ENUM_TIMEFRAMES tf;       // Source timeframe
};

struct VelocityData
{
   double   speedPipsPerBar;    // Average pips moved per bar
   int      direction;          // 1=UP, -1=DOWN, 0=NEUTRAL
   int      consecutiveBars;    // Consecutive same-direction bars
   double   score;              // Velocity score 0-100
   bool     confirmed;          // Velocity strong enough to trade
};

struct TradeState
{
   bool     inTrade;
   ulong    ticket;
   bool     tp1Hit;
   bool     breakevenSet;
   bool     trailingActive;
   double   entryPrice;
   double   initialSL;
   double   initialTP1;
   double   initialTP2;
   int      direction;          // 1=BUY, -1=SELL
};

//+------------------------------------------------------------------+
//| GLOBAL VARIABLES                                                  |
+------------------------------------------------------------------+

SRZone     Zones[];
int        TotalZones = 0;
TradeState CurrentTrade;
double     PipSize;
double     PipValue;
int        ATR_Handle;
string     CurrentZoneType = "NONE";  // "SUPPORT", "RESISTANCE", "NONE"
int        NearestZoneIdx  = -1;
bool       ZoneAlertSent   = false;
string     LogFilePath;

// Dashboard label names
string     DashPrefix = "HSR_";

//+------------------------------------------------------------------+
//| EXPERT INITIALIZATION                                             |
+------------------------------------------------------------------+
int OnInit()
{
   // Set pip size based on symbol digits
   PipSize = SymbolInfoDouble(_Symbol, SYMBOL_POINT);
   int digits = (int)SymbolInfoInteger(_Symbol, SYMBOL_DIGITS);
   if(digits == 3 || digits == 5) PipSize *= 10;
   if(StringFind(_Symbol, "JPY") >= 0 && digits == 3) PipSize = 0.01;

   // ATR indicator
   ATR_Handle = iATR(_Symbol, PERIOD_H1, (int)ATR_Period);
   if(ATR_Handle == INVALID_HANDLE)
   {
      Print("ERROR: Could not create ATR handle");
      return INIT_FAILED;
   }

   // Init trade state
   ResetTradeState();

   // Magic number and slippage
   Trade.SetExpertMagicNumber(MagicNumber);
   Trade.SetDeviationInPoints(MaxSlippage);

   // Log file path
   LogFilePath = TerminalInfoString(TERMINAL_DATA_PATH) + "\MQL5\Files\" + LogFileName;

   // Write CSV header if logging
   if(LogToFile) WriteLogHeader();

   // Build all S/R zones from history
   Print("=== HistoricalSR EA v2.0 Starting ===");
   Print("Building S/R Zones from full history...");
   BuildAllZones();
   Print("Total qualified zones found: ", TotalZones);

   // Draw dashboard
   if(ShowDashboard) DrawDashboard();

   return INIT_SUCCEEDED;
}

//+------------------------------------------------------------------+
//| EXPERT DEINITIALIZATION                                           |
+------------------------------------------------------------------+
void OnDeinit(const int reason)
{
   if(ATR_Handle != INVALID_HANDLE)
      IndicatorRelease(ATR_Handle);

   // Remove dashboard objects
   if(ShowDashboard) RemoveDashboard();
   ArrayFree(Zones);
}

//+------------------------------------------------------------------+
//| EXPERT TICK                                                       |
+------------------------------------------------------------------+
void OnTick()
{
   // Only process on new bar
   static datetime lastBar = 0;
   datetime currentBar = iTime(_Symbol, PERIOD_CURRENT, 0);
   if(currentBar == lastBar) 
   {
      // Still manage open trades on every tick
      if(CurrentTrade.inTrade) ManageOpenTrade();
      return;
   }
   lastBar = currentBar;

   // --- Step 1: Refresh zone data (recalculate strength) ---
   RefreshZoneActivity();

   // --- Step 2: Check if current price is near any zone ---
   double currentPrice = (iClose(_Symbol, PERIOD_CURRENT, 1) + iOpen(_Symbol, PERIOD_CURRENT, 1)) / 2.0;
   NearestZoneIdx = FindNearestActiveZone(currentPrice);

   // --- Step 3: Zone hit alert ---
   if(NearestZoneIdx >= 0 && AlertOnZoneHit && !ZoneAlertSent)
   {
      string zoneMsg = StringFormat("Price in %s Zone | Strength: %d | Price: %.5f",
         Zones[NearestZoneIdx].isResistance ? "RESISTANCE" : "SUPPORT",
         Zones[NearestZoneIdx].strength,
         Zones[NearestZoneIdx].price);
      Alert(_Symbol + " — " + zoneMsg);
      ZoneAlertSent = true;
   }
   else if(NearestZoneIdx < 0)
   {
      ZoneAlertSent = false;
   }

   // --- Step 4: Calculate velocity ---
   VelocityData vel = CalculateVelocity();

   // --- Step 5: Trade logic (only if no open trade) ---
   if(!CurrentTrade.inTrade)
   {
      if(NearestZoneIdx >= 0 && vel.confirmed)
      {
         EvaluateTradeEntry(NearestZoneIdx, vel);
      }
   }
   else
   {
      // Manage existing trade
      ManageOpenTrade();
   }

   // --- Step 6: Update dashboard ---
   if(ShowDashboard) UpdateDashboard(vel);
}

//+------------------------------------------------------------------+
//| BUILD ALL S/R ZONES FROM HISTORY (H4 + D1)                      |
+------------------------------------------------------------------+
void BuildAllZones()
{
   ArrayFree(Zones);
   TotalZones = 0;

   // Process H4 timeframe
   ProcessTimeframeZones(PERIOD_H4, H4_Bars);

   // Process D1 timeframe
   ProcessTimeframeZones(PERIOD_D1, D1_Bars);

   // Sort zones by strength descending
   SortZonesByStrength();

   Print("Zone building complete. H4+D1 zones: ", TotalZones);
}

//+------------------------------------------------------------------+
//| PROCESS ONE TIMEFRAME FOR ZONES                                   |
+------------------------------------------------------------------+
void ProcessTimeframeZones(ENUM_TIMEFRAMES tf, int barsToAnalyze)
{
   int available = Bars(_Symbol, tf);
   if(barsToAnalyze <= 0 || barsToAnalyze > available)
      barsToAnalyze = available - 1;

   if(barsToAnalyze < 10) return;

   double highArr[], lowArr[], closeArr[];
   long   volumeArr[];
   datetime timeArr[];

   if(CopyHigh(_Symbol, tf, 0, barsToAnalyze, highArr)  <= 0) return;
   if(CopyLow(_Symbol, tf, 0, barsToAnalyze, lowArr)    <= 0) return;
   if(CopyClose(_Symbol, tf, 0, barsToAnalyze, closeArr) <= 0) return;
   if(CopyTickVolume(_Symbol, tf, 0, barsToAnalyze, volumeArr) <= 0) return;
   if(CopyTime(_Symbol, tf, 0, barsToAnalyze, timeArr)  <= 0) return;

   int n = ArraySize(highArr);

   // Step 1: Find swing highs and lows
   double swingHighs[], swingLows[];
   datetime swingHighTimes[], swingLowTimes[];
   int swingHighCount = 0, swingLowCount = 0;
   int lb = SwingLookback;

   for(int i = lb; i < n - lb; i++)
   {
      bool isSwingHigh = true, isSwingLow = true;
      for(int j = 1; j <= lb; j++)
      {
         if(highArr[i] <= highArr[i-j] || highArr[i] <= highArr[i+j]) isSwingHigh = false;
         if(lowArr[i]  >= lowArr[i-j]  || lowArr[i]  >= lowArr[i+j])  isSwingLow  = false;
      }
      if(isSwingHigh)
      {
         ArrayResize(swingHighs, swingHighCount+1);
         ArrayResize(swingHighTimes, swingHighCount+1);
         swingHighs[swingHighCount] = highArr[i];
         swingHighTimes[swingHighCount] = timeArr[i];
         swingHighCount++;
      }
      if(isSwingLow)
      {
         ArrayResize(swingLows, swingLowCount+1);
         ArrayResize(swingLowTimes, swingLowCount+1);
         swingLows[swingLowCount] = lowArr[i];
         swingLowTimes[swingLowCount] = timeArr[i];
         swingLowCount++;
      }
   }

   // Step 2: Calculate average volume for normalization
   double avgVol = 0;
   for(int i = 0; i < n; i++) avgVol += (double)volumeArr[i];
   avgVol /= n;

   // Step 3: Build candidate price levels from swings
   // For each swing, count touches and sum volume in zone
   double zoneWidth = ZoneWidthPips * PipSize;

   // Combine all swing prices into candidate levels
   double candidates[];
   bool   candIsRes[];
   datetime candTimes[];
   int candCount = 0;

   for(int i = 0; i < swingHighCount; i++)
   {
      ArrayResize(candidates, candCount+1);
      ArrayResize(candIsRes, candCount+1);
      ArrayResize(candTimes, candCount+1);
      candidates[candCount] = swingHighs[i];
      candIsRes[candCount]  = true;
      candTimes[candCount]  = swingHighTimes[i];
      candCount++;
   }
   for(int i = 0; i < swingLowCount; i++)
   {
      ArrayResize(candidates, candCount+1);
      ArrayResize(candIsRes, candCount+1);
      ArrayResize(candTimes, candCount+1);
      candidates[candCount] = swingLows[i];
      candIsRes[candCount]  = false;
      candTimes[candCount]  = swingLowTimes[i];
      candCount++;
   }

   // Step 4: For each candidate, count touches and compute scores
   bool processed[];
   ArrayResize(processed, candCount);
   ArrayInitialize(processed, false);

   for(int c = 0; c < candCount; c++)
   {
      if(processed[c]) continue;

      double lvl = candidates[c];
      int    touchCount   = 0;
      double totalVolume  = 0;
      datetime lastTouch  = 0;
      bool   isSwingLevel = true;

      // Merge nearby candidates into this zone
      for(int c2 = c; c2 < candCount; c2++)
      {
         if(MathAbs(candidates[c2] - lvl) <= zoneWidth)
         {
            processed[c2] = true;
         }
      }

      // Count actual bar touches in price history
      for(int i = 0; i < n; i++)
      {
         bool touched = false;
         if(highArr[i] >= lvl - zoneWidth && lowArr[i] <= lvl + zoneWidth)
            touched = true;

         if(touched)
         {
            touchCount++;
            totalVolume += (double)volumeArr[i];
            if(timeArr[i] > lastTouch) lastTouch = timeArr[i];
         }
      }

      if(touchCount < MinTouchCount) continue;

      // Compute composite strength score
      double touchScore  = MathMin(100.0, (touchCount / 10.0) * 100.0);
      double volScore    = MathMin(100.0, (totalVolume / (touchCount * avgVol)) * 50.0);
      double swingScore  = isSwingLevel ? 30.0 : 0.0;

      // Recency bonus: zones touched more recently get higher score
      int barsAgo = (int)((TimeCurrent() - lastTouch) / PeriodSeconds(tf));
      double recencyScore = MathMax(0, 20.0 - (barsAgo / 50.0));

      int compositeScore = (int)(touchScore * 0.40 + volScore * 0.30 + swingScore * 0.20 + recencyScore * 0.10);
      compositeScore = MathMin(100, MathMax(0, compositeScore));

      if(compositeScore < ZoneStrengthMin) continue;

      // Add to global zones array
      ArrayResize(Zones, TotalZones + 1);
      Zones[TotalZones].price        = lvl;
      Zones[TotalZones].upperBound   = lvl + zoneWidth;
      Zones[TotalZones].lowerBound   = lvl - zoneWidth;
      Zones[TotalZones].touchCount   = touchCount;
      Zones[TotalZones].volumeScore  = volScore;
      Zones[TotalZones].isSwing      = isSwingLevel;
      Zones[TotalZones].strength     = compositeScore;
      Zones[TotalZones].isResistance = candIsRes[c];
      Zones[TotalZones].lastTouched  = lastTouch;
      Zones[TotalZones].active       = true;
      Zones[TotalZones].tf           = tf;
      TotalZones++;
   }
}

//+------------------------------------------------------------------+
//| SORT ZONES BY STRENGTH                                            |
+------------------------------------------------------------------+
void SortZonesByStrength()
{
   for(int i = 0; i < TotalZones - 1; i++)
      for(int j = i + 1; j < TotalZones; j++)
         if(Zones[j].strength > Zones[i].strength)
         {
            SRZone tmp  = Zones[i];
            Zones[i]    = Zones[j];
            Zones[j]    = tmp;
         }
}

//+------------------------------------------------------------------+
//| REFRESH ZONE ACTIVITY (deactivate broken zones)                  |
+------------------------------------------------------------------+
void RefreshZoneActivity()
{
   double close1 = iClose(_Symbol, PERIOD_CURRENT, 1);
   double close2 = iClose(_Symbol, PERIOD_CURRENT, 2);

   for(int i = 0; i < TotalZones; i++)
   {
      if(!Zones[i].active) continue;

      // Deactivate if price has cleanly closed beyond zone twice
      bool brokenUp   = (close1 > Zones[i].upperBound && close2 > Zones[i].upperBound);
      bool brokenDown = (close1 < Zones[i].lowerBound && close2 < Zones[i].lowerBound);

      if(Zones[i].isResistance && brokenUp)   Zones[i].active = false;
      if(!Zones[i].isResistance && brokenDown) Zones[i].active = false;
   }
}

//+------------------------------------------------------------------+
//| FIND NEAREST ACTIVE ZONE TO CURRENT PRICE                        |
+------------------------------------------------------------------+
int FindNearestActiveZone(double price)
{
   double minDist = DBL_MAX;
   int    nearest = -1;

   for(int i = 0; i < TotalZones; i++)
   {
      if(!Zones[i].active) continue;
      if(price >= Zones[i].lowerBound && price <= Zones[i].upperBound)
      {
         double dist = MathAbs(price - Zones[i].price);
         if(dist < minDist)
         {
            minDist = dist;
            nearest = i;
         }
      }
   }
   return nearest;
}

//+------------------------------------------------------------------+
//| CALCULATE VELOCITY                                                |
+------------------------------------------------------------------+
VelocityData CalculateVelocity()
{
   VelocityData vel;
   vel.speedPipsPerBar  = 0;
   vel.direction        = 0;
   vel.consecutiveBars  = 0;
   vel.score            = 0;
   vel.confirmed        = false;

   int bars = VelocityBars + 1;
   double opens[], closes[], highs[], lows[];
   if(CopyOpen(_Symbol, PERIOD_CURRENT, 1, bars, opens)   <= 0) return vel;
   if(CopyClose(_Symbol, PERIOD_CURRENT, 1, bars, closes) <= 0) return vel;
   if(CopyHigh(_Symbol, PERIOD_CURRENT, 1, bars, highs)   <= 0) return vel;
   if(CopyLow(_Symbol, PERIOD_CURRENT, 1, bars, lows)     <= 0) return vel;

   int n = MathMin(bars, ArraySize(closes));

   // Calculate average candle body size
   double totalBody = 0;
   int    upBars = 0, downBars = 0;
   for(int i = 0; i < n; i++)
   {
      totalBody += MathAbs(closes[i] - opens[i]);
      if(closes[i] > opens[i]) upBars++;
      else if(closes[i] < opens[i]) downBars++;
   }

   double avgBody    = totalBody / n;
   double speedPips  = avgBody / PipSize;
   vel.speedPipsPerBar = speedPips;

   // Determine direction
   if(upBars > downBars)   vel.direction = 1;
   else if(downBars > upBars) vel.direction = -1;
   else vel.direction = 0;

   // Count consecutive bars in current direction
   int streak = 0;
   for(int i = 0; i < n; i++)
   {
      bool isUp   = (closes[i] > opens[i]);
      bool isDown = (closes[i] < opens[i]);
      if(vel.direction == 1  && isUp)   streak++;
      else if(vel.direction == -1 && isDown) streak++;
      else break;
   }
   vel.consecutiveBars = streak;

   // Overall price movement over window
   double netMove = MathAbs(closes[0] - closes[n-1]) / PipSize;

   // Velocity score
   double speedScore   = MathMin(50.0, speedPips * 10.0);
   double streakScore  = MathMin(30.0, streak * 10.0);
   double netScore     = MathMin(20.0, netMove / 5.0);
   vel.score = speedScore + streakScore + netScore;

   // Confirmation
   vel.confirmed = (speedPips >= VelocityMinPips && streak >= ConsecutiveBarsMin);

   return vel;
}

//+------------------------------------------------------------------+
//| EVALUATE TRADE ENTRY                                             |
+------------------------------------------------------------------+
void EvaluateTradeEntry(int zoneIdx, VelocityData &vel)
{
   SRZone zone = Zones[zoneIdx];
   double currentPrice = SymbolInfoDouble(_Symbol, SYMBOL_BID);
   double atr = GetATR();
   if(atr <= 0) return;

   // Determine if price is bouncing or breaking through
   bool priceFalling = (vel.direction == -1);
   bool priceRising  = (vel.direction == 1);

   int    tradeDir  = 0;
   string tradeType = "";

   if(zone.isResistance)
   {
      // At resistance: expect bounce DOWN or breakout UP
      double breakThreshold = zone.upperBound + (ZoneWidthPips * BreakoutThresholdPct / 100.0) * PipSize;

      if(priceRising && currentPrice > breakThreshold)
      {
         // BREAKOUT UP through resistance → BUY
         tradeDir  = 1;
         tradeType = "BREAKOUT";
      }
      else if(priceFalling)
      {
         // BOUNCE at resistance → SELL
         tradeDir  = -1;
         tradeType = "BOUNCE";
      }
   }
   else
   {
      // At support: expect bounce UP or breakout DOWN
      double breakThreshold = zone.lowerBound - (ZoneWidthPips * BreakoutThresholdPct / 100.0) * PipSize;

      if(priceFalling && currentPrice < breakThreshold)
      {
         // BREAKOUT DOWN through support → SELL
         tradeDir  = -1;
         tradeType = "BREAKOUT";
      }
      else if(priceRising)
      {
         // BOUNCE at support → BUY
         tradeDir  = 1;
         tradeType = "BOUNCE";
      }
   }

   if(tradeDir == 0) return;

   // Confirmation: check last N candles confirm direction
   if(!ConfirmDirection(tradeDir)) return;

   // Calculate SL and TP
   double sl, tp1, tp2;
   double slDistance = atr * ATR_SL_Multiplier;

   if(tradeDir == 1)
   {
      double ask = SymbolInfoDouble(_Symbol, SYMBOL_ASK);
      sl  = ask - slDistance;
      tp1 = ask + slDistance * TP1_RR;
      tp2 = ask + slDistance * TP2_RR;

      // SL just below zone
      if(tradeType == "BOUNCE")
         sl = MathMin(sl, zone.lowerBound - 2 * PipSize);
   }
   else
   {
      double bid = SymbolInfoDouble(_Symbol, SYMBOL_BID);
      sl  = bid + slDistance;
      tp1 = bid - slDistance * TP1_RR;
      tp2 = bid - slDistance * TP2_RR;

      // SL just above zone
      if(tradeType == "BOUNCE")
         sl = MathMax(sl, zone.upperBound + 2 * PipSize);
   }

   // Calculate lot size
   double lots = CalculateLotSize(slDistance);
   if(lots <= 0) return;

   // Place order
   bool result = false;
   string comment = StringFormat("HSR|%s|Z%d|S%d", tradeType, zoneIdx, zone.strength);

   if(tradeDir == 1)
      result = Trade.Buy(lots, _Symbol, 0, sl, tp2, comment);
   else
      result = Trade.Sell(lots, _Symbol, 0, sl, tp2, comment);

   if(result)
   {
      ulong ticket = Trade.ResultOrder();
      CurrentTrade.inTrade      = true;
      CurrentTrade.ticket       = ticket;
      CurrentTrade.tp1Hit       = false;
      CurrentTrade.breakevenSet = false;
      CurrentTrade.trailingActive = false;
      CurrentTrade.direction    = tradeDir;
      CurrentTrade.entryPrice   = (tradeDir == 1) ?
         SymbolInfoDouble(_Symbol, SYMBOL_ASK) :
         SymbolInfoDouble(_Symbol, SYMBOL_BID);
      CurrentTrade.initialSL    = sl;
      CurrentTrade.initialTP1   = tp1;
      CurrentTrade.initialTP2   = tp2;

      string msg = StringFormat(
         "TRADE OPEN | %s %s | %s | Zone: %.5f | Strength: %d | Lots: %.2f | SL: %.5f | TP1: %.5f | TP2: %.5f",
         (tradeDir == 1 ? "BUY" : "SELL"), _Symbol, tradeType,
         zone.price, zone.strength, lots, sl, tp1, tp2);

      Print(msg);
      if(SendEmailOnTrade) SendMail("HSR EA Trade Open — " + _Symbol, msg);
      if(SendPushOnTrade)  SendNotification(msg);
      if(LogToFile)        WriteTradeLog("OPEN", ticket, tradeDir, lots, CurrentTrade.entryPrice, sl, tp1, tp2, zone.strength, tradeType);

      // Deactivate zone after trade (prevent re-entry)
      Zones[zoneIdx].active = false;
   }
   else
   {
      Print("Trade failed: ", Trade.ResultRetcodeDescription());
   }
}

//+------------------------------------------------------------------+
//| CONFIRM DIRECTION WITH LAST N CANDLES                            |
+------------------------------------------------------------------+
bool ConfirmDirection(int direction)
{
   int needed = ConfirmationBars;
   double opens[], closes[];
   if(CopyOpen(_Symbol, PERIOD_CURRENT, 1, needed, opens)   <= 0) return false;
   if(CopyClose(_Symbol, PERIOD_CURRENT, 1, needed, closes) <= 0) return false;

   int confirmed = 0;
   for(int i = 0; i < needed; i++)
   {
      if(direction == 1  && closes[i] > opens[i]) confirmed++;
      if(direction == -1 && closes[i] < opens[i]) confirmed++;
   }
   return (confirmed >= needed);
}

//+------------------------------------------------------------------+
//| MANAGE OPEN TRADE (Breakeven, Trailing, Partial Close)           |
+------------------------------------------------------------------+
void ManageOpenTrade()
{
   if(!CurrentTrade.inTrade) return;

   // Check if position still exists
   if(!PositionSelectByTicket(CurrentTrade.ticket))
   {
      // Position closed (TP/SL hit)
      if(LogToFile) WriteTradeLog("CLOSED", CurrentTrade.ticket, CurrentTrade.direction,
         0, 0, 0, 0, 0, 0, "AUTO");
      ResetTradeState();
      return;
   }

   double currentPrice = (CurrentTrade.direction == 1) ?
      SymbolInfoDouble(_Symbol, SYMBOL_BID) :
      SymbolInfoDouble(_Symbol, SYMBOL_ASK);

   double entryPrice = CurrentTrade.entryPrice;
   double currentSL  = PositionGetDouble(POSITION_SL);
   double pipProfit  = (currentPrice - entryPrice) * CurrentTrade.direction / PipSize;

   // --- Partial Close at TP1 ---
   if(!CurrentTrade.tp1Hit)
   {
      bool tp1Reached = (CurrentTrade.direction == 1 && currentPrice >= CurrentTrade.initialTP1) ||
                        (CurrentTrade.direction == -1 && currentPrice <= CurrentTrade.initialTP1);

      if(tp1Reached)
      {
         double currentLots = PositionGetDouble(POSITION_VOLUME);
         double lotStep     = SymbolInfoDouble(_Symbol, SYMBOL_VOLUME_STEP);
         double minLot      = SymbolInfoDouble(_Symbol, SYMBOL_VOLUME_MIN);
         double rawClose    = currentLots * (PartialClosePercent / 100.0);
         double closeLots   = MathFloor(rawClose / lotStep) * lotStep;
         closeLots = NormalizeDouble(closeLots, 2);
         closeLots = MathMax(closeLots, minLot);

         if(CurrentTrade.direction == 1)
            Trade.Sell(closeLots, _Symbol, 0, 0, 0, "HSR|PartialClose");
         else
            Trade.Buy(closeLots, _Symbol, 0, 0, 0, "HSR|PartialClose");

         CurrentTrade.tp1Hit = true;
         Print("Partial close executed at TP1: ", closeLots, " lots");
      }
   }

   // --- Move SL to Breakeven ---
   if(!CurrentTrade.breakevenSet && pipProfit >= BreakevenPips)
   {
      double newSL = entryPrice + (CurrentTrade.direction * 1 * PipSize); // 1 pip above/below entry
      if(CurrentTrade.direction == 1 && newSL > currentSL)
      {
         Trade.PositionModify(CurrentTrade.ticket, newSL, CurrentTrade.initialTP2);
         CurrentTrade.breakevenSet = true;
         Print("SL moved to breakeven: ", newSL);
      }
      else if(CurrentTrade.direction == -1 && newSL < currentSL)
      {
         Trade.PositionModify(CurrentTrade.ticket, newSL, CurrentTrade.initialTP2);
         CurrentTrade.breakevenSet = true;
         Print("SL moved to breakeven: ", newSL);
      }
   }

   // --- Trailing Stop ---
   if(pipProfit >= TrailingStartPips)
   {
      CurrentTrade.trailingActive = true;
      double trailDistance = TrailingStepPips * PipSize;
      double newSL;

      if(CurrentTrade.direction == 1)
      {
         newSL = currentPrice - trailDistance;
         if(newSL > currentSL)
            Trade.PositionModify(CurrentTrade.ticket, newSL, CurrentTrade.initialTP2);
      }
      else
      {
         newSL = currentPrice + trailDistance;
         if(newSL < currentSL)
            Trade.PositionModify(CurrentTrade.ticket, newSL, CurrentTrade.initialTP2);
      }
   }
}

//+------------------------------------------------------------------+
//| CALCULATE LOT SIZE BASED ON RISK %                               |
+------------------------------------------------------------------+
double CalculateLotSize(double slDistance)
{
   double balance    = AccountInfoDouble(ACCOUNT_BALANCE);
   double riskAmount = balance * RiskPercent / 100.0;

   double tickValue  = SymbolInfoDouble(_Symbol, SYMBOL_TRADE_TICK_VALUE);
   double tickSize   = SymbolInfoDouble(_Symbol, SYMBOL_TRADE_TICK_SIZE);
   double minLot     = SymbolInfoDouble(_Symbol, SYMBOL_VOLUME_MIN);
   double maxLot     = SymbolInfoDouble(_Symbol, SYMBOL_VOLUME_MAX);
   double lotStep    = SymbolInfoDouble(_Symbol, SYMBOL_VOLUME_STEP);

   if(tickSize == 0 || tickValue == 0) return minLot;

   double slTicks = slDistance / tickSize;
   double lots    = riskAmount / (slTicks * tickValue);

   lots = MathFloor(lots / lotStep) * lotStep;
   lots = MathMax(minLot, MathMin(maxLot, lots));

   return NormalizeDouble(lots, 2);
}

//+------------------------------------------------------------------+
//| GET ATR VALUE                                                     |
+------------------------------------------------------------------+
double GetATR()
{
   double atrBuf[];
   if(CopyBuffer(ATR_Handle, 0, 1, 1, atrBuf) <= 0) return 0;
   return atrBuf[0];
}

//+------------------------------------------------------------------+
//| RESET TRADE STATE                                                 |
+------------------------------------------------------------------+
void ResetTradeState()
{
   CurrentTrade.inTrade        = false;
   CurrentTrade.ticket         = 0;
   CurrentTrade.tp1Hit         = false;
   CurrentTrade.breakevenSet   = false;
   CurrentTrade.trailingActive = false;
   CurrentTrade.entryPrice     = 0;
   CurrentTrade.initialSL      = 0;
   CurrentTrade.initialTP1     = 0;
   CurrentTrade.initialTP2     = 0;
   CurrentTrade.direction      = 0;
}

//+------------------------------------------------------------------+
//| DRAW ON-CHART DASHBOARD                                          |
+------------------------------------------------------------------+
void DrawDashboard()
{
   string prefix = DashPrefix;
   int x = 10, y = 20;
   int lineH = 18;
   color bgColor   = C'20,20,30';
   color borderClr = C'80,120,200';
   color titleClr  = C'120,180,255';
   color textClr   = clrSilver;

   // Background rectangle
   ObjectCreate(0, prefix+"BG", OBJ_RECTANGLE_LABEL, 0, 0, 0);
   ObjectSetInteger(0, prefix+"BG", OBJPROP_XDISTANCE, x-5);
   ObjectSetInteger(0, prefix+"BG", OBJPROP_YDISTANCE, y-5);
   ObjectSetInteger(0, prefix+"BG", OBJPROP_XSIZE, 320);
   ObjectSetInteger(0, prefix+"BG", OBJPROP_YSIZE, 260);
   ObjectSetInteger(0, prefix+"BG", OBJPROP_BGCOLOR, bgColor);
   ObjectSetInteger(0, prefix+"BG", OBJPROP_BORDER_COLOR, borderClr);
   ObjectSetInteger(0, prefix+"BG", OBJPROP_BORDER_TYPE, BORDER_FLAT);
   ObjectSetInteger(0, prefix+"BG", OBJPROP_CORNER, CORNER_LEFT_UPPER);

   // Title
   CreateLabel(prefix+"T0", "⬛ HistoricalSR EA v2.0", x+5, y+2, titleClr, 10, true);
   CreateLabel(prefix+"T1", "──────────────────────────────", x+5, y+18, borderClr, 8, false);

   // Labels (will be updated dynamically)
   CreateLabel(prefix+"L_symbol",  "Symbol  : " + _Symbol,  x+5, y+32,  textClr, 9, false);
   CreateLabel(prefix+"L_zones",   "Zones   : 0 active",    x+5, y+50,  textClr, 9, false);
   CreateLabel(prefix+"L_zone",    "Zone    : ---",          x+5, y+68,  textClr, 9, false);
   CreateLabel(prefix+"L_zstr",    "Strength: ---",          x+5, y+86,  textClr, 9, false);
   CreateLabel(prefix+"L_vel",     "Velocity: ---",          x+5, y+104, textClr, 9, false);
   CreateLabel(prefix+"L_dir",     "Direction: ---",         x+5, y+122, textClr, 9, false);
   CreateLabel(prefix+"L_trade",   "Trade   : NO TRADE",     x+5, y+140, textClr, 9, false);
   CreateLabel(prefix+"L_be",      "Breakeven: ---",         x+5, y+158, textClr, 9, false);
   CreateLabel(prefix+"L_trail",   "Trailing: ---",          x+5, y+176, textClr, 9, false);
   CreateLabel(prefix+"L_risk",    "Risk    : 1.0%",         x+5, y+194, textClr, 9, false);
   CreateLabel(prefix+"L_tf",      "TF      : H4 + D1",     x+5, y+212, textClr, 9, false);
   CreateLabel(prefix+"L_sep2",    "──────────────────────────────", x+5, y+226, borderClr, 8, false);
   CreateLabel(prefix+"L_copy",    "Forex | Gold | Crypto | Indices", x+5, y+236, C'60,80,120', 8, false);

   ChartRedraw();
}

//+------------------------------------------------------------------+
//| CREATE A LABEL OBJECT                                            |
+------------------------------------------------------------------+
void CreateLabel(string name, string text, int x, int y, color clr, int fontSize, bool bold)
{
   ObjectCreate(0, name, OBJ_LABEL, 0, 0, 0);
   ObjectSetInteger(0, name, OBJPROP_XDISTANCE, x);
   ObjectSetInteger(0, name, OBJPROP_YDISTANCE, y);
   ObjectSetInteger(0, name, OBJPROP_CORNER, CORNER_LEFT_UPPER);
   ObjectSetString(0, name, OBJPROP_TEXT, text);
   ObjectSetInteger(0, name, OBJPROP_COLOR, clr);
   ObjectSetInteger(0, name, OBJPROP_FONTSIZE, fontSize);
   ObjectSetString(0, name, OBJPROP_FONT, bold ? "Arial Bold" : "Arial");
   ObjectSetInteger(0, name, OBJPROP_SELECTABLE, false);
}

//+------------------------------------------------------------------+
//| UPDATE DASHBOARD                                                  |
+------------------------------------------------------------------+
void UpdateDashboard(VelocityData &vel)
{
   string prefix = DashPrefix;

   // Count active zones
   int activeZones = 0;
   for(int i = 0; i < TotalZones; i++)
      if(Zones[i].active) activeZones++;

   ObjectSetString(0, prefix+"L_zones", OBJPROP_TEXT,
      StringFormat("Zones   : %d active / %d total", activeZones, TotalZones));

   if(NearestZoneIdx >= 0)
   {
      SRZone z = Zones[NearestZoneIdx];
      string zType = z.isResistance ? "RESISTANCE" : "SUPPORT";
      string tfStr = (z.tf == PERIOD_H4 ? "H4" : "D1");

      ObjectSetString(0, prefix+"L_zone", OBJPROP_TEXT,
         StringFormat("Zone    : %s [%s] @ %.5f", zType, tfStr, z.price));
      ObjectSetInteger(0, prefix+"L_zone", OBJPROP_COLOR,
         z.isResistance ? clrOrangeRed : clrLimeGreen);

      ObjectSetString(0, prefix+"L_zstr", OBJPROP_TEXT,
         StringFormat("Strength: %d/100 | Touches: %d", z.strength, z.touchCount));
      ObjectSetInteger(0, prefix+"L_zstr", OBJPROP_COLOR,
         z.strength >= 70 ? clrGold : z.strength >= 50 ? clrYellow : clrSilver);
   }
   else
   {
      ObjectSetString(0, prefix+"L_zone", OBJPROP_TEXT, "Zone    : Not in zone");
      ObjectSetInteger(0, prefix+"L_zone", OBJPROP_COLOR, clrGray);
      ObjectSetString(0, prefix+"L_zstr", OBJPROP_TEXT, "Strength: ---");
   }

   // Velocity
   string dirStr = vel.direction == 1 ? "▲ UP" : vel.direction == -1 ? "▼ DOWN" : "─ FLAT";
   color  dirClr = vel.direction == 1 ? clrLimeGreen : vel.direction == -1 ? clrOrangeRed : clrGray;
   ObjectSetString(0, prefix+"L_vel", OBJPROP_TEXT,
      StringFormat("Velocity: %.1f pip/bar | Score: %.0f", vel.speedPipsPerBar, vel.score));
   ObjectSetString(0, prefix+"L_dir", OBJPROP_TEXT,
      StringFormat("Direction: %s | Streak: %d bars | %s",
         dirStr, vel.consecutiveBars, vel.confirmed ? "✔ CONFIRMED" : "✘ WEAK"));
   ObjectSetInteger(0, prefix+"L_dir", OBJPROP_COLOR, vel.confirmed ? dirClr : clrGray);

   // Trade state
   if(CurrentTrade.inTrade)
   {
      string tDir = (CurrentTrade.direction == 1) ? "BUY  ▲" : "SELL ▼";
      color  tClr = (CurrentTrade.direction == 1) ? clrLimeGreen : clrOrangeRed;
      ObjectSetString(0, prefix+"L_trade", OBJPROP_TEXT,
         StringFormat("Trade   : %s | Ticket: %d", tDir, (int)CurrentTrade.ticket));
      ObjectSetInteger(0, prefix+"L_trade", OBJPROP_COLOR, tClr);

      ObjectSetString(0, prefix+"L_be", OBJPROP_TEXT,
         StringFormat("Breakeven: %s", CurrentTrade.breakevenSet ? "✔ SET" : "Pending"));
      ObjectSetInteger(0, prefix+"L_be", OBJPROP_COLOR,
         CurrentTrade.breakevenSet ? clrGold : clrGray);

      ObjectSetString(0, prefix+"L_trail", OBJPROP_TEXT,
         StringFormat("Trailing: %s", CurrentTrade.trailingActive ? "✔ ACTIVE" : "Pending"));
      ObjectSetInteger(0, prefix+"L_trail", OBJPROP_COLOR,
         CurrentTrade.trailingActive ? clrGold : clrGray);
   }
   else
   {
      ObjectSetString(0, prefix+"L_trade", OBJPROP_TEXT, "Trade   : WAITING FOR SETUP");
      ObjectSetInteger(0, prefix+"L_trade", OBJPROP_COLOR, clrGray);
      ObjectSetString(0, prefix+"L_be",    OBJPROP_TEXT, "Breakeven: ---");
      ObjectSetString(0, prefix+"L_trail", OBJPROP_TEXT, "Trailing: ---");
   }

   ChartRedraw();
}

//+------------------------------------------------------------------+
//| REMOVE DASHBOARD                                                  |
+------------------------------------------------------------------+
void RemoveDashboard()
{
   string prefix = DashPrefix;
   string labels[] = {"BG","T0","T1","L_symbol","L_zones","L_zone","L_zstr",
                       "L_vel","L_dir","L_trade","L_be","L_trail","L_risk","L_tf","L_sep2","L_copy"};
   for(int i = 0; i < ArraySize(labels); i++)
      ObjectDelete(0, prefix + labels[i]);
}

//+------------------------------------------------------------------+
//| WRITE LOG HEADER                                                  |
+------------------------------------------------------------------+
void WriteLogHeader()
{
   int handle = FileOpen(LogFileName, FILE_WRITE|FILE_CSV|FILE_ANSI, ',');
   if(handle == INVALID_HANDLE) return;
   FileWrite(handle, "DateTime","Symbol","Action","Direction","Type","Lots",
             "EntryPrice","SL","TP1","TP2","ZoneStrength","Ticket");
   FileClose(handle);
}

//+------------------------------------------------------------------+
//| WRITE TRADE LOG ENTRY                                            |
+------------------------------------------------------------------+
void WriteTradeLog(string action, ulong ticket, int dir, double lots,
                   double entry, double sl, double tp1, double tp2,
                   int strength, string tradeType)
{
   int handle = FileOpen(LogFileName, FILE_READ|FILE_WRITE|FILE_CSV|FILE_ANSI, ',');
   if(handle == INVALID_HANDLE) return;
   FileSeek(handle, 0, SEEK_END);
   FileWrite(handle,
      TimeToString(TimeCurrent(), TIME_DATE|TIME_SECONDS),
      _Symbol, action,
      (dir == 1 ? "BUY" : "SELL"),
      tradeType,
      DoubleToString(lots, 2),
      DoubleToString(entry, _Digits),
      DoubleToString(sl, _Digits),
      DoubleToString(tp1, _Digits),
      DoubleToString(tp2, _Digits),
      IntegerToString(strength),
      IntegerToString((int)ticket));
   FileClose(handle);
}

//+------------------------------------------------------------------+
//| END OF EA                                                         |
+------------------------------------------------------------------+
