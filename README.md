# Capstone-Project-Data-Cleaning-and-Analysis-with-Spreadsheet/ Excel

## Q.1) What was the city with the highest sales?

![image](https://user-images.githubusercontent.com/116772724/222916659-46a14019-4169-4647-a48f-67e6c4c81d12.png)


## Q.2) What is the average discount given for all orders?

=AVERAGE(N2:N9995)
![image](https://user-images.githubusercontent.com/116772724/222921229-4cbc91c8-24ef-47d9-8ade-69f6ef710901.png)


## Q.3) What is the most popular product among customers in the "Consumer" segment?

![image](https://user-images.githubusercontent.com/116772724/222916583-52498b6c-ab5e-415e-ae44-e5b59a8ab433.png)

## Q.4) What is the total profit made for the "Office Supplies" category?

![image](https://user-images.githubusercontent.com/116772724/222916787-20d94bd8-8ab4-4c6a-8bfd-7f7af8f2b099.png)

## Q.5) Who is the customer who has made the most purchases? (Hint: use the “Order ID column to answer the question.)

![image](https://user-images.githubusercontent.com/116772724/222916881-7834a08f-e5ba-4fcc-be7b-edf00d2d2608.png)

## Q.6) What state made the most profit?

![image](https://user-images.githubusercontent.com/116772724/222917238-61a7f9f6-a1fa-4f32-b7f8-8ead00ef07d1.png)

## Q.7) How many orders were shipped via "Standard Class" ship mode?

![image](https://user-images.githubusercontent.com/116772724/222917412-db69379a-befa-46d9-a840-941b7763ddc3.png)

## Q.8) Which region had the highest sales in the month of June?

![image](https://user-images.githubusercontent.com/116772724/222918849-e3298812-b812-495e-83d2-c24919d775c1.png)

## Q.9) Calculate the price per unit of each product (before discounts), and put it in a separate column. What's the most expensive product?
 
 ![image](https://user-images.githubusercontent.com/116772724/222920556-b197f502-fdb1-497b-8e3e-395b0182ef73.png)

## Q.10) Create a pivot table that shows the total sales for each manufacturer and category combination. In the "Technology" category, which manufacturer had the second highest sales?
 
 ![image](https://user-images.githubusercontent.com/116772724/222920934-16e22695-2904-4500-8c3c-4e7db913fd92.png)

## Q.11)  Create a new column that calculates the profit margin for each order (hint: profit/sales). What's the profit margin average?

=AVERAGE(R2:R9995)
![image](https://user-images.githubusercontent.com/116772724/222921199-23aa0313-e386-4665-8920-5d143395e549.png)

## Q.12) Use a VLOOKUP function to create a new column that shows the product sub-category for each product based on the separate sub-category sheet.What is the subcategory of “Xerox 1887”?

![image](https://user-images.githubusercontent.com/116772724/222922445-ba2770d5-264e-47e2-84df-320289eb0984.png)

## Q.13) Create a new column that calculates the number of days between the order date and the ship date for each order. 
  ##     Create a conditional formatting “color scale” for this column, from greenish to reddish.
  ##     What is the number of days for order id - “CA-2015-100363”?
  
  ![image](https://user-images.githubusercontent.com/116772724/222931688-d1f75c51-066e-40c3-9c5c-7d05c2aac7bc.png)
  
## Q.14) Use the INDEX and MATCH functions to create a new column that shows the shipping cost for each order based on the separate shipping prices sheet. Assume that quantity or weight doesn’t matter. What is the shipping price for order id “CA-2015-100678”?

![image](https://user-images.githubusercontent.com/116772724/222988329-16982fc3-f051-482e-a0d4-18115ca16b2d.png)


## Q.15) Create a new column that concatenates the customer name, city, and state into a single string for each order. Select the correct result for CA-2015-100090

![image](https://user-images.githubusercontent.com/116772724/222939003-331490ec-3367-4c5c-ac98-4dae4b4fc197.png)

## Q.16)  Use the IFS function to create a new column that categorizes each order as "High," "Low," or "Loss" based on profit and sales criteria.
      "High" consider as:

   - If sales are above 200 and profit is above 20

    - If profit is above 40.

       Else:

 If the profit is equal or below 0 this is categorized as “Loss”
Any other case this is categorized as "Low"
Use conditional formatting to color the columns with the values “High” in green and the value “Loss” in red.**

How many “Loss” do you have?

![image](https://user-images.githubusercontent.com/116772724/222988530-e960190f-b77a-4354-8023-da05d70cd02d.png)

## Q.17) In a new sheet, create a dropdown of category and product which returns the price for a unit (which you previously solved in exercise 9.)

![image](https://user-images.githubusercontent.com/116772724/222988588-4b402594-d74c-4450-bf66-ae59276086d2.png)
