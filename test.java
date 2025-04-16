class pubg{

   pubg(){
       
    System.out.println("This is a constructor");

    }

    static void Something(){

        System.out.println("Some works");
    }
}


class test{

    public static void main(String[] args) {
        
       //constructor - it is a special type of function it calls automatically when the object is created
       //constructor name is same as class name
       pubg game = new pubg();

    }
}