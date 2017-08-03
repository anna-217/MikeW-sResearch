package gui.swingx;

import javax.swing.*;

/**
 * Created by aaaaaqi5 on 8/2/2017.
 */
public class JavaSwingExample {
    public static void main(String[] args) {
        javax.swing.SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                createAndShowGUI();
            }
        });

    }

    private static void createAndShowGUI() {
        //Make sure we have nice window decorations.

        JFrame.setDefaultLookAndFeelDecorated(true);

        //Create and set up the window.

        JFrame frame = new JFrame("JavaSwingExample");

        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        //Add the ubiquitous "Hello World" label.

        JLabel label = new JLabel("Hello World");

        frame.getContentPane().add(label);

        //Display the window.

        frame.pack();

        frame.setVisible(true);

    }

}
