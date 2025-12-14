# LiftSheet 
[中文](README.zh.md)

## What is LiftSheet?



LiftSheet is a powerful, highly modern **template/spreadsheet for strength training**, designed to be the ultimate **training programming and data tracking system** for coaches and athletes.

It does not provide preset programs, but serves as your **central hub for training program management**, allowing you to:

- Access scientifically comprehensive and periodized training program design sheets.
- Utilize multi-dimensional statistical charts to fully quantify and track training progress and performance.
- Benefit from a highly automated weight calculator/regulator for dynamic, real-time training content.
- Integrate multiple powerful calculators and include a built-in sidebar AI chat function.

Say goodbye to primitive notes and crude spreadsheets! LiftSheet is the ultimate tool for transcending traditional methods and achieving efficient strength training management.

![2](assets/2.png)

![20250527j728a](assets/20250527j728a.png)

## **Why be LiftSheet? For what?**

Strength sports resemble car racing: success hinges on “how to build a faster car.” Continuous progress requires optimizing quantified data and making consistent, data-driven adjustments tailored to the individual.

**LiftSheet is not a specific product; it’s an idea.** Its core purpose is to inspire focused athletes to build their own training programming and data tracking system. Rather than settling for the current offering, **LiftSheet aims to be the spark that enables you to DIY the tool best suited for your needs.**

- **For Coaches/Program Writers:** Use this toolset to structure your content. The program's content and its structure are mutually reinforcing; a powerful tool expands our perspective on the content itself.
- **For Athletes with Coaches**: Import your coach's programs into your own management system. Even with a coach, taking responsibility for your own progress is crucial. Enhancing self-awareness and understanding through data is key to long-term, stable progress, rather than solely relying on the coach for problem-solving.



## **Get LiftSheet Now**



**Access Link:** https://docs.google.com/spreadsheets/d/1NL3w-Peuepj49wIi83sbfOzNWiHEbX3nnhz_MV23wzE/edit?usp=sharing

Type the link into your browser or click the link directly.

![image-20251121022109400](assets/image-20251121022109400.png)

Select **"Make a copy"** to save a personal copy to your Google Sheets library, and you can start using it happily!

<img src="assets/image-20251121022232203.png" alt="image-20251121022232203" style="zoom: 67%;" />

<img src="assets/image-20251121022443117.png" alt="image-20251121022443117" style="zoom:67%;" />



## **Read This Scenario, and You'll Know Exactly How to Use It!**

### **1. Writing the Program**

You are Coach Taylor, and you are now setting up the training plan for your athlete Susan for **M1B1W1D1** (Mesocycle 1 - Block 1 - Week 1 - Day 1).

First, you confirm that Susan's M1B1 will use a **4-day-per-week** training frequency. You therefore locate the `🗓️ M1B1 (4Days)` template among the three frequency options. Next, you right-click on `🗓️ M1B1 (3Days)` and `🗓️ M1B1 (5Days)`and hide them.

<img src="assets/image-20251122220119111.png" alt="image-20251122220119111" style="zoom:50%;" />

You open `🗓️ M1B1 (4Days)` and set the date for W1D1 to **Tuesday, 2025/08/05**. Next, you set the **Gap (days)** to **2**. The spreadsheet automatically adjusts the date of the next training day to **Thursday, 2025/08/07**, a two-day interval.

<img src="assets/image-20251122220022563.png" alt="image-20251122220022563" style="zoom:50%;" />

Now it's time to input the program: The main lift, **Competition Squat**, is programmed as: **Top single@RPE7**; **Working Set: 4x5@RPE6**.

<img src="assets/image-20251121000554728.png" alt="image-20251121000554728" style="zoom:50%;" />

Next, you need to calculate the target weight for Susan's 1-rep set at RPE 7. You open the `🔢 calculator`, enter an Estimated 1RM of **120**, **RPE 7**, and **1 rep** into the RPE Weight Calculator, which returns **107.5** (kg).

<img src="assets/image-20251121000642939.png" alt="image-20251121000642939" style="zoom:50%;" />

You input **107.5** into the blank **Target Load** cell in the first row of W1D1. The weights for the working sets are then automatically calculated and populated based on the program content. Your planning work is nearly complete; the columns on the right are reserved for Susan to fill in after her training session.

<img src="assets/image-20251121000911113.png" alt="image-20251121000911113" style="zoom:50%;" />

You also observe that the spreadsheet automatically tracks key metrics such as the total number of **sets**, **volume**, **this week's max weight**, and **estimated 1RM**. The **Delta** column shows the difference compared to the previous week (which will be blank for the first week).

<img src="assets/image-20251121000957619.png" alt="image-20251121000957619" style="zoom:50%;" />

<img src="assets/image-20251121001033334.png" alt="image-20251121001033334" style="zoom:50%;" />



****



### 2.Execution and Logging 

Susan starts by logging her weight in the **📈Weight Tracker**.

<img src="assets/image-20251121001336815.png" alt="image-20251121001336815" style="zoom:50%;" />

She uses the **Warm Up Calculator** (Target load 107.5 1rep RPE 7) to get **5 suggested warm-up sets**.

<img src="assets/image-20251121001413212.png" alt="image-20251121001413212" style="zoom:50%;" />

She notes that if she aims for RPE 9 at 115kg, the calculator provides the answer.

<img src="assets/image-20251121001531688.png" alt="image-20251121001531688" style="zoom:50%;" />

If her RPE 7 attempt feels like RPE 8, she changes the input (e.g., 107.5 Attempt RPE to 8). The calculator dynamically suggests a new RPE 9 Load (e.g., from 115kg down to **110kg**). Susan finds this feature highly intelligent.

<img src="assets/image-20251121001858671.png" alt="image-20251121001858671" style="zoom:50%;" />

During W1D1, due to poor readiness, Susan uses **105kg** for her top set (Actual RPE 7). 于是在Actual PRE中选择了Optimal（或者手动选择RPE数字）

<img src="assets/image-20251121002036945.png" alt="image-20251121002036945" style="zoom:50%;" />

The table automatically adjusts: **Target load drops** (e.g., 92.5 to 90), and **Actual Reps** are auto-filled (or manually entered for numerical RPE). Susan only needs to select the **Actual RPE** after each set to complete the log. 

<img src="assets/image-20251121002036945-3738846.png" alt="image-20251121002036945" style="zoom:50%;" />

The **statistics module updates in real-time.**

<img src="assets/image-20251121002242598.png" alt="image-20251121002242598" style="zoom:50%;" />



****



### 3.Statistical Analysis of Training Data

Susan has completed today's squat training. Now, let's try to perform a statistical analysis of the data.

<img src="assets/image-20251121003132631.png" alt="image-20251121003132631" style="zoom:50%;" />

To analyze the data, open the **📈Performance Tracker**. **Do not manually copy data.**

<img src="assets/image-20251121003411656.png" alt="image-20251121003411656" style="zoom:50%;" />

Use the **"MyScript"** button (next to Help in the menu bar) and select **Sync Performance Data**.

<img src="assets/image-20251121003736363.png" alt="image-20251121003736363" style="zoom:50%;" />

Input the target Block (e.g., **M1B1**) or the full sheet name (e.g., **🗓️M1B1 (4Days)**).

<img src="assets/image-20251121003910992.png" alt="image-20251121003910992" style="zoom:50%;" />

If multiple sheets are found, specify the correct number (**e.g., 2**).

<img src="assets/image-20251121004116326.png" alt="image-20251121004116326" style="zoom:50%;" />

After synchronization, **M1B1 data and charts appear** in the **📈Performance Tracker**. 

<img src="assets/image-20251121004156004.png" alt="image-20251121004156004" style="zoom:50%;" />

Since it uses linked cells, **future updates to M1B1 are automatically reflected.**

<img src="assets/image-20251121004319934.png" alt="image-20251121004319934" style="zoom:50%;" />

The data is also visible in the **📊Volume Tracker**.

<img src="assets/image-20251121004600514.png" alt="image-20251121004600514" style="zoom:50%;" />



****



### 4.Customizing Your Program

#### 4.1Add Your Favorite Exercise

Click the **pen edit icon**.

<img src="assets/image-20251121005738480.png" alt="image-20251121005738480" style="zoom:50%;" />

In the side panel, select **Add another item**, fill in the details, and select **Done, Apply to all.**

<img src="assets/image-20251121005849107.png" alt="image-20251121005849107" style="zoom:50%;" />

#### 4.2Creating a New Block

New Blocks (e.g., M1B2) are created by **copying one of the template sheets** (**🗓️M1B1 (3Days/4Days/5Days)**) to retain all functions.

<img src="assets/image-20251121005023846.png" alt="image-20251121005023846" style="zoom:50%;" />

#### 4.3Customizing the Weekly Training Structure

Copy a template sheet based on your frequency. 

<img src="assets/image-20251121010814407.png" alt="image-20251121010814407" style="zoom:50%;" />

Weekly content is divided into **Main, Secondary Main, and Accessory lifts**, which can be rearranged.

<img src="assets/image-20251121011157834.png" alt="image-20251121011157834" style="zoom:50%;" />

**Google Sheets Feature:** Cutting/dragging cells within the same sheet **dynamically adapts cell referencing.**

**Example:** To swap Bench Press (Day 1 Secondary) with Deadlift (Day 2 Secondary), **cut** the Bench Press block elsewhere

<img src="assets/image-20251121012447384.png" alt="image-20251121012447384" style="zoom:50%;" />

Then **cut and paste** the Deadlift block under the Day 1 Squat.

<img src="assets/image-20251121012700999.png" alt="image-20251121012700999" style="zoom:50%;" />

Rearrange the data in the bottom statistics section (using cut/drag) to match the new structure.

![image-20251121014212592](assets/image-20251121014212592.png)

![image-20251121012817198](assets/image-20251121012817198.png)

![image-20251121013350562](assets/image-20251121013350562.png)

Once Week 1 is set, **copy the entire structure to Week 2 and Week 3.**



****



### 5.Advanced Usage

#### 5.1Sidebar AI Chat

**Configure API key** (DeepSeek API).

![image-20251121020123372](assets/image-20251121020123372.png)

Select **Open Chat Window**.

![image-20251121020339488](assets/image-20251121020339488.png)

The AI can **analyze training data** and **plan future arrangements** by directly accessing the sheet's data—solving a limitation of traditional AI tools.

#### 5.2Plate Loading Calculator

![image-20251121014902222](assets/image-20251121014902222.png)

**Left Side (Weight Calculation):** Input plates (25-5kg for large, 2.5-0.5kg for small) and Barbell/Collar weight to find the **Total Weight** (e.g., 152.5kg).

**Right Side (Plate Selection):** Input Barbell weight and **Target Weight** (e.g., 170kg) to get the **required plates per side**(e.g., 3 x 25kg).

#### 5.3Powerlifting IPF GL Points Calculator

![image-20251121015842815](assets/image-20251121015842815.png)

#### 5.4Taper Calculator

![image-20251121015944876](assets/image-20251121015944876.png)
